import time
import threading
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from pathlib import Path
import pythoncom

class VBAFileHandler(FileSystemEventHandler):
    def __init__(self, importer, watcher, extensions=['.bas', '.cls', '.frm'], debug=False, sync_delete_modules=False):
        self.importer = importer
        self.watcher = watcher
        self.extensions = extensions
        self.last_modified = {}
        self.debug = debug
        self.sync_delete_modules = sync_delete_modules
    
    def on_modified(self, event):
        if event.is_directory:
            return
        file_path = Path(event.src_path)
        if file_path.suffix not in self.extensions:
            return
        if self.watcher.is_exporting:
            if self.debug:
                print(f"[DEBUG] Ignoring change during export: {file_path.name}")
            return
        # FIX: Ignore empty file for import
        if file_path.stat().st_size < 10:
            if self.debug:
                print(f"[DEBUG] Ignoring empty file for import: {file_path.name}")
            return
        current_time = time.time()
        last_time = self.last_modified.get(str(file_path), 0)
        if current_time - last_time < 1.0:
            if self.debug:
                print(f"[DEBUG] Debouncing: {file_path.name}")
            return
        self.last_modified[str(file_path)] = current_time
        try:
            rel_path = file_path.relative_to(self.watcher.watch_directory)
            print(f"\n📝 Change detected: {rel_path}")
        except ValueError:
            print(f"\n📝 Change detected: {file_path.name}")
        
        # The import_module method now handles COM initialization internally
        self.importer.import_module(file_path)

    def on_deleted(self, event):
        if not self.sync_delete_modules or event.is_directory:
            return
        
        # COM initialization for this thread
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            if self.debug:
                print(f"[DEBUG] COM initialized for on_deleted handler")
        except:
            if self.debug:
                print(f"[DEBUG] COM already initialized for on_deleted handler")
            pass
        
        try:
            from .vba_import import VisioVBAImporter
            
            file_path = Path(event.src_path)
            if file_path.suffix.lower() not in self.extensions:
                return
            
            # FIX: Only remove module if module existed and was not empty
            module_name = file_path.stem
            importer_threadlocal = VisioVBAImporter(self.importer.visio_file_path, debug=self.debug)
            if not importer_threadlocal.connect_to_visio():
                print("⚠️  Could not connect to Visio for module removal.")
                return
            
            for doc_info in importer_threadlocal.doc_manager.get_all_documents_with_vba():
                vb_project = doc_info.doc.VBProject
                for comp in vb_project.VBComponents:
                    if comp.Name == module_name and comp.CodeModule.CountOfLines > 0:
                        try:
                            vb_project.VBComponents.Remove(comp)
                            print(f"✓ Removed Visio module: {module_name} ({doc_info.name})")
                            if self.debug:
                                print(f"[DEBUG] Module '{module_name}' removed from '{doc_info.name}' due to local delete")
                        except Exception as e:
                            print(f"⚠️  Error removing module '{module_name}' from '{doc_info.name}': {e}")
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] Error in on_deleted handler: {e}")
                import traceback
                traceback.print_exc()
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                    if self.debug:
                        print(f"[DEBUG] COM uninitialized for on_deleted handler")
                except:
                    pass

class VBAWatcher:
    def __init__(self, watch_directory, importer, exporter=None, bidirectional=False, debug=False, sync_delete_modules=False):
        self.watch_directory = watch_directory
        self.importer = importer
        self.exporter = exporter
        self.bidirectional = bidirectional
        self.debug = debug
        self.observer = None
        self.smart_poll_timer = None
        self.last_vba_sync_time = 0
        self.last_export_hashes = {}  # Track hash per document: {doc_folder: hash}
        self.is_exporting = False  # Flag to prevent concurrent operations
        self.doc = importer.doc
        self.sync_delete_modules = sync_delete_modules
    
    def _pause_observer(self):
        if self.observer and self.observer.is_alive():
            if self.debug:
                print("[DEBUG] Pausing observer...")
            self.observer.stop()
            self.observer.join(timeout=2)
    
    def _resume_observer(self):
        if self.observer and not self.observer.is_alive():
            if self.debug:
                print("[DEBUG] Restarting observer...")
            event_handler = VBAFileHandler(self.importer, self, debug=self.debug, sync_delete_modules=self.sync_delete_modules)
            self.observer = Observer()
            self.observer.schedule(
                event_handler,
                str(self.watch_directory),
                recursive=True  # Watch subdirectories for multi-document support
            )
            self.observer.start()
    
    def _start_polling(self, poll_interval=4):
        self.smart_poll_timer = threading.Timer(poll_interval, self._poll_vba_changes)
        self.smart_poll_timer.daemon = True
        self.smart_poll_timer.start()
    
    def _poll_vba_changes(self):
        """Poll for VBA changes in Visio and export to VS Code.
        
        This runs in a separate thread, so COM must be initialized.
        """
        # COM initialization for this thread
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            if self.debug:
                print("[DEBUG] COM initialized for polling thread")
        except:
            if self.debug:
                print("[DEBUG] COM already initialized for polling thread")
            pass
        
        try:
            from .vba_import import VisioVBAImporter
            from .vba_export import VisioVBAExporter
            
            local_importer = VisioVBAImporter(getattr(self.importer, 'visio_file_path', None), debug=self.debug)
            if not local_importer.connect_to_visio():
                if self.debug:
                    print("[DEBUG] Reconnection failed, waiting for next cycle...")
                return
            if self.debug:
                print("[DEBUG] Connection established successfully in poll thread")
            
            if self.exporter:
                self.is_exporting = True
                self._pause_observer()
                try:
                    thread_exporter = VisioVBAExporter(str(local_importer.visio_file_path), debug=self.debug)
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
                finally:
                    time.sleep(0.5)
                    self._resume_observer()
                    self.is_exporting = False
            
            self.last_vba_sync_time = time.time()
        except Exception as e:
            print(f"⚠️  Error during polling export: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            self.is_exporting = False
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                    if self.debug:
                        print("[DEBUG] COM uninitialized for polling thread")
                except:
                    pass
            
            if self.bidirectional:
                self._start_polling()
    
    def start(self):
        event_handler = VBAFileHandler(self.importer, self, debug=self.debug, sync_delete_modules=self.sync_delete_modules)
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
        
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            self.stop()
    
    def stop(self):
        if self.observer:
            self.observer.stop()
            self.observer.join()
            print("\n✓ Monitoring stopped")
        if self.smart_poll_timer:
            self.smart_poll_timer.cancel()
            print("✓ Polling stopped")