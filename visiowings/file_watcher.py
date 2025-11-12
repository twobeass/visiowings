"""File system watcher for VBA modules
Adds smart polling for VSCode ← Visio sync (detects document changes)
"""
import time
import threading
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from pathlib import Path

class VBAFileHandler(FileSystemEventHandler):
    def __init__(self, importer, extensions=['.bas', '.cls', '.frm']):
        self.importer = importer
        self.extensions = extensions
        self.last_modified = {}
    def on_modified(self, event):
        if event.is_directory:
            return
        file_path = Path(event.src_path)
        if file_path.suffix not in self.extensions:
            return
        current_time = time.time()
        last_time = self.last_modified.get(str(file_path), 0)
        if current_time - last_time < 1.0:
            return
        self.last_modified[str(file_path)] = current_time
        print(f"\n📝 Änderung erkannt: {file_path.name}")
        self.importer.import_module(file_path)

class VBAWatcher:
    def __init__(self, watch_directory, importer, exporter=None, bidirectional=False):
        self.watch_directory = watch_directory
        self.importer = importer
        self.exporter = exporter
        self.bidirectional = bidirectional
        self.observer = None
        self.smart_poll_timer = None
        self.last_vba_sync_time = 0
        self.doc = importer.doc
    def _start_polling(self, poll_interval=4):
        """Check periodically for changes in Visio (sync → VSCode)"""
        self.smart_poll_timer = threading.Timer(poll_interval, self._poll_vba_changes)
        self.smart_poll_timer.daemon = True
        self.smart_poll_timer.start()
    def _poll_vba_changes(self):
        try:
            # Reconnect in case document has been lost
            if not self.importer._ensure_connection():
                return
            # Simple sync: always export - could be optimized with last doc hash
            if self.exporter:
                self.exporter.export_modules(self.watch_directory)
                print("🔄 Visio-Dokument wurde synchronisiert → VSCode.")
            self.last_vba_sync_time = time.time()
        except Exception as e:
            print(f"⚠️  Fehler beim Polling-Export: {e}")
        finally:
            # Continue polling if enabled
            if self.bidirectional:
                self._start_polling()
    def start(self):
        event_handler = VBAFileHandler(self.importer)
        self.observer = Observer()
        self.observer.schedule(
            event_handler,
            str(self.watch_directory),
            recursive=False
        )
        self.observer.start()
        print(f"\n👁️  Überwache Verzeichnis: {self.watch_directory}")
        print("💾 Speichere Dateien in VS Code (Ctrl+S) um sie nach Visio zu synchronisieren")
        print("⏸️  Drücke Ctrl+C zum Beenden...\n")
        if self.bidirectional and self.exporter:
            print("🔄 Bidirektionaler Sync: Änderungen in Visio werden automatisch nach VSCode exportiert.")
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
            print("\n✓ Überwachung beendet")
        # Polling stoppen
        if self.smart_poll_timer:
            self.smart_poll_timer.cancel()
            print("✓ Polling gestoppt")
