import signal
import sys
import threading
import time
from collections import OrderedDict
from pathlib import Path

import pythoncom
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer


# Watchdog dispatches events from a worker thread, so we have to guard the
# debounce dict against concurrent reads/writes. A small LRU cap also avoids
# unbounded memory growth in long-running sessions.
_DEBOUNCE_MAX_ENTRIES = 1000
_DEBOUNCE_WINDOW_SECONDS = 1.0


class VBAFileHandler(FileSystemEventHandler):
    def __init__(self, importer, watcher, extensions=None, debug=False, sync_delete_modules=False):
        if extensions is None:
            extensions = ['.bas', '.cls', '.frm']
        self.importer = importer
        self.watcher = watcher
        self.extensions = extensions
        self._debounce_lock = threading.Lock()
        self._last_modified: "OrderedDict[str, float]" = OrderedDict()
        self.debug = debug
        self.sync_delete_modules = sync_delete_modules

    def _record_change(self, key: str, now: float) -> bool:
        """Return True if this event should be processed, False if debounced.

        Implements a thread-safe sliding-window debounce with an LRU cap to
        prevent the dict from growing without bound when many files churn.
        """

        with self._debounce_lock:
            last = self._last_modified.get(key, 0.0)
            if now - last < _DEBOUNCE_WINDOW_SECONDS:
                return False
            self._last_modified[key] = now
            self._last_modified.move_to_end(key)
            while len(self._last_modified) > _DEBOUNCE_MAX_ENTRIES:
                self._last_modified.popitem(last=False)
            return True

    def on_created(self, event):
        """Handle creation of new files"""
        self._handle_change(event, change_type="created")

    def on_modified(self, event):
        """Handle modification of existing files"""
        self._handle_change(event, change_type="modified")

    def _handle_change(self, event, change_type="modified"):
        # Check shutdown flag
        if self.watcher.shutdown_requested:
            return

        if event.is_directory:
            return

        file_path = Path(event.src_path)

        if file_path.suffix not in self.extensions:
            return

        if self.watcher.is_exporting:
            if self.debug:
                print(f"[DEBUG] Ignoring {change_type} during export: {file_path.name}")
            return

        # Validate file exists and is not empty
        try:
            if not file_path.exists() or file_path.stat().st_size < 10:
                if self.debug:
                    print(f"[DEBUG] Ignoring empty or non-existent file: {file_path.name}")
                return
        except OSError as e:
            if self.debug:
                print(f"[DEBUG] Error checking file: {type(e).__name__}: {e}")
            return

        # Debounce rapid changes (thread-safe).
        if not self._record_change(str(file_path), time.time()):
            if self.debug:
                print(f"[DEBUG] Debouncing: {file_path.name}")
            return

        try:
            rel_path = file_path.relative_to(self.watcher.watch_directory)
            print(f"\n📝 Change detected ({change_type}): {rel_path}")
        except ValueError:
            print(f"\n📝 Change detected ({change_type}): {file_path.name}")

        # Import module (handles COM initialization internally)
        try:
            self.importer.import_module(file_path, edit_mode=True)

        except Exception as e:
            print(f"❌ Error during import ({type(e).__name__}): {e}")
            if self.debug:
                import traceback
                traceback.print_exc()

    def on_deleted(self, event):
        # Check shutdown flag
        if self.watcher.shutdown_requested:
            return

        if not self.sync_delete_modules or event.is_directory:
            return

        # COM initialization for this thread
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            if self.debug:
                print("[DEBUG] COM initialized for on_deleted handler")
        except pythoncom.com_error as e:
            if self.debug:
                print(f"[DEBUG] COM already initialized for on_deleted handler: {e}")

        try:
            from .vba_import import VisioVBAImporter

            file_path = Path(event.src_path)
            if file_path.suffix.lower() not in self.extensions:
                return

            module_name = file_path.stem
            use_rubberduck = getattr(self.importer, 'use_rubberduck', False)
            importer_threadlocal = VisioVBAImporter(self.importer.visio_file_path, debug=self.debug, use_rubberduck=use_rubberduck)

            if not importer_threadlocal.connect_to_visio():
                print("⚠️  Could not connect to Visio for module removal.")
                return

            found_module = False
            for doc_info in importer_threadlocal.doc_manager.get_all_documents_with_vba():
                try:
                    vb_project = doc_info.doc.VBProject
                    for comp in vb_project.VBComponents:
                        if comp.Name == module_name:
                            found_module = True
                            # Only remove if module has content
                            if comp.CodeModule.CountOfLines > 0:
                                try:
                                    vb_project.VBComponents.Remove(comp)
                                    print(f"✓ Removed Visio module: {module_name} ({doc_info.name})")
                                    if self.debug:
                                        print(f"[DEBUG] Module '{module_name}' removed from '{doc_info.name}' due to local delete")
                                except (AttributeError, pythoncom.com_error) as e:
                                    print(f"⚠️  Error removing module '{module_name}' from '{doc_info.name}': {type(e).__name__}: {e}")
                            break
                except (AttributeError, pythoncom.com_error) as e:
                    if self.debug:
                        print(f"[DEBUG] Error accessing document {doc_info.name}: {type(e).__name__}: {e}")

            if not found_module and self.debug:
                print(f"[DEBUG] Module '{module_name}' not found in any Visio document")

        except Exception as e:
            if self.debug:
                print(f"[DEBUG] Error in on_deleted handler: {type(e).__name__}: {e}")
                import traceback
                traceback.print_exc()
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                    if self.debug:
                        print("[DEBUG] COM uninitialized for on_deleted handler")
                except pythoncom.com_error as e:
                    if self.debug:
                        print(f"[DEBUG] Error uninitializing COM: {e}")


class VBAWatcher:
    def __init__(self, watch_directory, importer, exporter=None, bidirectional=False, debug=False, sync_delete_modules=False):
        self.watch_directory = watch_directory
        self.importer = importer
        self.exporter = exporter
        self.bidirectional = bidirectional
        self.debug = debug
        self.observer = None

        # Synchronisation primitives.
        # _state_lock guards observer/timer mutation across threads.
        # _exporting is a thread-safe flag the watchdog handler reads to
        # ignore self-induced events while we push exports to disk.
        # _shutdown is set on signal/stop and checked everywhere we might
        # otherwise schedule new work.
        self._state_lock = threading.RLock()
        self._exporting = threading.Event()
        self._shutdown = threading.Event()

        self.smart_poll_timer = None
        self.last_vba_sync_time = 0
        self.last_export_hashes = {}  # Track hash per document: {doc_folder: hash}
        self.doc = importer.doc
        self.sync_delete_modules = sync_delete_modules

    # ----- backwards-compatible attribute aliases ---------------------- #
    @property
    def is_exporting(self) -> bool:
        return self._exporting.is_set()

    @is_exporting.setter
    def is_exporting(self, value: bool) -> None:
        if value:
            self._exporting.set()
        else:
            self._exporting.clear()

    @property
    def shutdown_requested(self) -> bool:
        return self._shutdown.is_set()

    @shutdown_requested.setter
    def shutdown_requested(self, value: bool) -> None:
        if value:
            self._shutdown.set()
        else:
            self._shutdown.clear()

    def _pause_observer(self):
        """Pause file system observer"""
        with self._state_lock:
            obs = self.observer
            if not (obs and obs.is_alive()):
                return
            if self.debug:
                print("[DEBUG] Pausing observer...")
            try:
                obs.stop()
                obs.join(timeout=3)
            except Exception as e:
                if self.debug:
                    print(f"[DEBUG] Error pausing observer: {e}")

    def _resume_observer(self):
        """Resume file system observer"""
        with self._state_lock:
            if self._shutdown.is_set():
                return

            if self.observer and not self.observer.is_alive():
                if self.debug:
                    print("[DEBUG] Restarting observer...")
                try:
                    event_handler = VBAFileHandler(self.importer, self, debug=self.debug, sync_delete_modules=self.sync_delete_modules)
                    self.observer = Observer()
                    self.observer.schedule(
                        event_handler,
                        str(self.watch_directory),
                        recursive=True  # Watch subdirectories for multi-document support
                    )
                    self.observer.start()
                except Exception as e:
                    print(f"⚠️  Error restarting observer: {e}")
                    if self.debug:
                        import traceback
                        traceback.print_exc()

    def _start_polling(self, poll_interval=4):
        """Start polling timer for bidirectional sync"""
        with self._state_lock:
            if self._shutdown.is_set():
                return
            # If a previous timer is still queued, cancel it first to keep
            # the schedule predictable.
            if self.smart_poll_timer is not None:
                try:
                    self.smart_poll_timer.cancel()
                except Exception:  # noqa: BLE001 - best effort
                    pass
            timer = threading.Timer(poll_interval, self._poll_vba_changes)
            timer.daemon = True
            self.smart_poll_timer = timer
            timer.start()

    def _poll_vba_changes(self):
        """Poll for VBA changes in Visio and export to VS Code.

        This runs in a separate thread, so COM must be initialized.
        """
        if self._shutdown.is_set():
            return

        # COM initialization for this thread
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            if self.debug:
                print("[DEBUG] COM initialized for polling thread")
        except pythoncom.com_error as e:
            if self.debug:
                print(f"[DEBUG] COM already initialized for polling thread: {e}")

        try:
            from .vba_export import VisioVBAExporter
            from .vba_import import VisioVBAImporter

            use_rubberduck = getattr(self.importer, 'use_rubberduck', False)
            local_importer = VisioVBAImporter(getattr(self.importer, 'visio_file_path', None), debug=self.debug, use_rubberduck=use_rubberduck)

            if not local_importer.connect_to_visio():
                if self.debug:
                    print("[DEBUG] Reconnection failed, waiting for next cycle...")
                return

            if self.debug:
                print("[DEBUG] Connection established successfully in poll thread")

            if self.exporter and not self._shutdown.is_set():
                self._exporting.set()
                self._pause_observer()

                try:
                    thread_exporter = VisioVBAExporter(str(local_importer.visio_file_path), debug=self.debug, use_rubberduck=use_rubberduck)

                    if thread_exporter.connect_to_visio(silent=True):
                        all_exported, all_hashes = thread_exporter.export_modules(
                            self.watch_directory,
                            last_hashes=self.last_export_hashes
                        )

                        if all_exported:
                            exported_count = sum(len(files) for files in all_exported.values())
                            if exported_count > 0:
                                self.last_export_hashes = all_hashes
                                if self.debug:
                                    print(f"[DEBUG] Hashes updated: {list(all_hashes.keys())}")
                                print("🔄 Visio document(s) synchronized → VS Code.")
                            elif self.debug:
                                print("[DEBUG] No changes detected in Visio, no export.")
                        else:
                            if all_hashes:
                                self.last_export_hashes = all_hashes
                            if self.debug:
                                print("[DEBUG] No changes detected in Visio, no export.")

                except Exception as e:
                    print(f"⚠️  Error during export: {e}")
                    if self.debug:
                        import traceback
                        traceback.print_exc()
                finally:
                    time.sleep(0.5)
                    self._resume_observer()
                    self._exporting.clear()

            self.last_vba_sync_time = time.time()

        except Exception as e:
            print(f"⚠️  Error during polling export: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            self._exporting.clear()
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                    if self.debug:
                        print("[DEBUG] COM uninitialized for polling thread")
                except pythoncom.com_error as e:
                    if self.debug:
                        print(f"[DEBUG] Error uninitializing COM: {e}")

            # Schedule next poll if bidirectional and not shutting down.
            # _start_polling re-checks shutdown under the lock, which closes
            # the race between stop() and a tail-end reschedule here.
            if self.bidirectional and not self._shutdown.is_set():
                self._start_polling()

    def _handle_shutdown(self, signum, frame):
        """Handle shutdown signals gracefully"""
        print("\n\n⏸️  Shutting down gracefully...")
        self._shutdown.set()
        self.stop()
        sys.exit(0)

    def start(self):
        """Start file watcher and optional bidirectional polling"""
        # Register signal handlers for graceful shutdown
        signal.signal(signal.SIGINT, self._handle_shutdown)
        if hasattr(signal, 'SIGTERM'):
            signal.signal(signal.SIGTERM, self._handle_shutdown)

        try:
            event_handler = VBAFileHandler(self.importer, self, debug=self.debug, sync_delete_modules=self.sync_delete_modules)
            with self._state_lock:
                self.observer = Observer()
                self.observer.schedule(
                    event_handler,
                    str(self.watch_directory),
                    recursive=True  # Watch subdirectories for multi-document support
                )
                self.observer.start()

            print(f"\n👁️  Watching directory: {self.watch_directory}")
            print("💾 Save files in VS Code (Ctrl+S) to synchronize them to Visio")
            print("⏸️  Press Ctrl+C to stop...\n")

            if self.bidirectional and self.exporter:
                print("🔄 Bidirectional sync: Changes in Visio are automatically exported to VS Code.")
                if self.debug:
                    print("[DEBUG] Debug mode enabled\n")
                self._start_polling()

            # Main loop
            while not self._shutdown.is_set():
                time.sleep(1)

        except KeyboardInterrupt:
            self._handle_shutdown(None, None)
        except Exception as e:
            print(f"\n❌ Watcher error: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            self.stop()

    def stop(self):
        """Stop watcher and clean up resources"""
        self._shutdown.set()

        with self._state_lock:
            timer = self.smart_poll_timer
            self.smart_poll_timer = None

        # Stop polling timer
        if timer is not None:
            try:
                timer.cancel()
                if self.debug:
                    print("[DEBUG] Polling timer cancelled")
            except Exception as e:
                if self.debug:
                    print(f"[DEBUG] Error cancelling timer: {e}")

        # Stop observer
        with self._state_lock:
            obs = self.observer
            self.observer = None

        if obs is not None:
            try:
                obs.stop()
                obs.join(timeout=5)
                if self.debug:
                    print("[DEBUG] Observer stopped")
            except Exception as e:
                if self.debug:
                    print(f"[DEBUG] Error stopping observer: {e}")

        print("✓ Monitoring stopped")
        if self.bidirectional:
            print("✓ Polling stopped")
