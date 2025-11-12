"""File system watcher for VBA modules
Adds smart polling for VSCode ← Visio sync (detects document changes).
Fix: COM threading (CoInitialize) in poll thread.
Fix: Hash-based change detection to prevent endless loops.
Fix: Pause observer during export to prevent file watcher triggers.
"""
import time
import threading
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from pathlib import Path

class VBAFileHandler(FileSystemEventHandler):
    def __init__(self, importer, watcher, extensions=['.bas', '.cls', '.frm'], debug=False):
        self.importer = importer
        self.watcher = watcher
        self.extensions = extensions
        self.last_modified = {}
        self.debug = debug
    
    def on_modified(self, event):
        if event.is_directory:
            return
        
        file_path = Path(event.src_path)
        if file_path.suffix not in self.extensions:
            return
        
        # Ignore if we're currently exporting
        if self.watcher.is_exporting:
            if self.debug:
                print(f"[DEBUG] Ignoriere Änderung während Export: {file_path.name}")
            return
        
        # Debouncing: ignore rapid successive changes
        current_time = time.time()
        last_time = self.last_modified.get(str(file_path), 0)
        if current_time - last_time < 1.0:
            if self.debug:
                print(f"[DEBUG] Debouncing: {file_path.name}")
            return
        
        self.last_modified[str(file_path)] = current_time
        
        print(f"\n📝 Änderung erkannt: {file_path.name}")
        self.importer.import_module(file_path)

class VBAWatcher:
    def __init__(self, watch_directory, importer, exporter=None, bidirectional=False, debug=False):
        self.watch_directory = watch_directory
        self.importer = importer
        self.exporter = exporter
        self.bidirectional = bidirectional
        self.debug = debug
        self.observer = None
        self.smart_poll_timer = None
        self.last_vba_sync_time = 0
        self.last_export_hash = None  # Track hash between polling cycles
        self.is_exporting = False  # Flag to prevent concurrent operations
        self.doc = importer.doc
    
    def _pause_observer(self):
        """Pause file system observer during export"""
        if self.observer and self.observer.is_alive():
            if self.debug:
                print("[DEBUG] Pausiere Observer...")
            self.observer.stop()
            self.observer.join(timeout=2)
    
    def _resume_observer(self):
        """Resume file system observer after export"""
        if self.observer and not self.observer.is_alive():
            if self.debug:
                print("[DEBUG] Starte Observer neu...")
            # Create new observer with same handler
            event_handler = VBAFileHandler(self.importer, self, debug=self.debug)
            self.observer = Observer()
            self.observer.schedule(
                event_handler,
                str(self.watch_directory),
                recursive=False
            )
            self.observer.start()
    
    def _start_polling(self, poll_interval=4):
        """Start polling timer"""
        self.smart_poll_timer = threading.Timer(poll_interval, self._poll_vba_changes)
        self.smart_poll_timer.daemon = True
        self.smart_poll_timer.start()
    
    def _check_connection_silent(self):
        """Check if connection is still active without verbose logging"""
        try:
            _ = self.importer.doc.Name
            return True
        except:
            return False
    
    def _poll_vba_changes(self):
        """Poll for VBA changes in Visio and export if changed"""
        try:
            import pythoncom
            pythoncom.CoInitialize()
            
            # Check connection silently first
            if not self._check_connection_silent():
                if self.debug:
                    print("[DEBUG] Verbindung verloren, versuche neu zu verbinden...")
                if not self.importer._ensure_connection():
                    if self.debug:
                        print("[DEBUG] Wiederverbindung fehlgeschlagen, warte auf nächsten Zyklus...")
                    return
                elif self.debug:
                    print("[DEBUG] Wiederverbindung erfolgreich")
            
            if self.exporter:
                # Set export flag to prevent file watcher from triggering
                self.is_exporting = True
                
                # Pause observer before export
                self._pause_observer()
                
                try:
                    # Create new exporter for this thread
                    from .vba_export import VisioVBAExporter
                    thread_exporter = VisioVBAExporter(
                        str(self.importer.visio_file_path), 
                        debug=self.debug
                    )
                    
                    # Connect silently (document is already open)
                    if thread_exporter.connect_to_visio():
                        # Export with hash comparison
                        result = thread_exporter.export_modules(
                            self.watch_directory, 
                            last_hash=self.last_export_hash
                        )
                        
                        if result and len(result) == 2:
                            exported_files, current_hash = result
                            
                            if exported_files:  # Files were actually exported (hash changed)
                                self.last_export_hash = current_hash
                                if self.debug:
                                    print(f"[DEBUG] Hash aktualisiert: {current_hash[:8]}...")
                                print("🔄 Visio-Dokument wurde synchronisiert → VSCode.")
                            else:
                                # No changes - update hash but don't export
                                self.last_export_hash = current_hash
                                if self.debug:
                                    print("[DEBUG] Keine Änderungen in Visio erkannt, kein Export.")
                        elif self.debug:
                            print("[DEBUG] Export-Result ungültig")
                
                finally:
                    # Always resume observer and clear flag
                    time.sleep(0.5)  # Short delay before resuming
                    self._resume_observer()
                    self.is_exporting = False
            
            self.last_vba_sync_time = time.time()
        
        except Exception as e:
            print(f"⚠️  Fehler beim Polling-Export: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            self.is_exporting = False
        
        finally:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except: 
                pass
            
            # Schedule next poll
            if self.bidirectional:
                self._start_polling()
    
    def start(self):
        """Start file watcher and optional polling"""
        event_handler = VBAFileHandler(self.importer, self, debug=self.debug)
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
            if self.debug:
                print("[DEBUG] Debug-Modus aktiviert\n")
            self._start_polling()
        
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            self.stop()
    
    def stop(self):
        """Stop file watcher and polling"""
        if self.observer:
            self.observer.stop()
            self.observer.join()
            print("\n✓ Überwachung beendet")
        
        if self.smart_poll_timer:
            self.smart_poll_timer.cancel()
            print("✓ Polling gestoppt")
