"""Command Line Interface for visiowings with bidirectional sync support
Now supports multiple documents (drawings + stencils)
"""
import argparse
from pathlib import Path
from .vba_export import VisioVBAExporter
from .vba_import import VisioVBAImporter
from .file_watcher import VBAWatcher

def cmd_edit(args):
    """Edit command: Export + Watch + Import with live sync"""
    visio_file = Path(args.file).resolve()
    output_dir = Path(args.output or '.').resolve()
    debug = getattr(args, 'debug', False)
    
    if not visio_file.exists():
        print(f"❌ Datei nicht gefunden: {visio_file}")
        return
    
    print(f"📂 Visio-Datei: {visio_file}")
    print(f"📁 Export-Verzeichnis: {output_dir}")
    if debug:
        print("[DEBUG] Debug-Modus aktiviert")
    
    print("\n=== Exportiere VBA-Module ===")
    exporter = VisioVBAExporter(str(visio_file), debug=debug)
    if not exporter.connect_to_visio():
        return
    
    # Export returns dict format for multi-document support
    all_exported, all_hashes = exporter.export_modules(output_dir)
    
    if not all_exported:
        print("❌ Keine Module exportiert")
        return
    
    # Count total exported files
    total_files = sum(len(files) for files in all_exported.values())
    total_docs = len(all_exported)
    
    if total_docs > 1:
        print(f"\n✓ {total_files} Module aus {total_docs} Dokumenten exportiert")
    else:
        print(f"\n✓ {total_files} Module exportiert")
    
    if debug:
        for doc_folder, doc_hash in all_hashes.items():
            print(f"[DEBUG] {doc_folder}: Hash {doc_hash[:8]}...")
    
    print("\n=== Starte Live-Synchronisation ===")
    importer = VisioVBAImporter(str(visio_file), force_document=args.force, debug=debug)
    if not importer.connect_to_visio():
        return
    
    watcher = VBAWatcher(
        output_dir, 
        importer, 
        exporter=exporter, 
        bidirectional=getattr(args, 'bidirectional', False),
        debug=debug
    )
    watcher.last_export_hashes = all_hashes  # Fix: Übernehme initialen Export-Hash zum Watcher
    watcher.start()

def cmd_export(args):
    """Export command: Export VBA modules only"""
    visio_file = Path(args.file).resolve()
    output_dir = Path(args.output or '.').resolve()
    debug = getattr(args, 'debug', False)
    
    exporter = VisioVBAExporter(str(visio_file), debug=debug)
    if exporter.connect_to_visio():
        all_exported, all_hashes = exporter.export_modules(output_dir)
        
        if all_exported:
            total_files = sum(len(files) for files in all_exported.values())
            total_docs = len(all_exported)
            
            if total_docs > 1:
                print(f"\n✓ {total_files} Module aus {total_docs} Dokumenten exportiert")
            else:
                print(f"\n✓ {total_files} Module exportiert")
            
            if debug:
                for doc_folder, doc_hash in all_hashes.items():
                    print(f"[DEBUG] {doc_folder}: Hash {doc_hash[:8]}...")

def cmd_import(args):
    """Import command: Import VBA modules only"""
    visio_file = Path(args.file).resolve()
    input_dir = Path(args.input or '.').resolve()
    debug = getattr(args, 'debug', False)
    
    importer = VisioVBAImporter(str(visio_file), force_document=args.force, debug=debug)
    if importer.connect_to_visio():
        imported_count = 0
        
        # Import from root directory (backward compatibility)
        for ext in ['*.bas', '*.cls', '*.frm']:
            for file in input_dir.glob(ext):
                if importer.import_module(file):
                    imported_count += 1
        
        # Import from subdirectories (multi-document support)
        for doc_folder in importer.get_document_folders():
            doc_dir = input_dir / doc_folder
            if doc_dir.exists() and doc_dir.is_dir():
                for ext in ['*.bas', '*.cls', '*.frm']:
                    for file in doc_dir.glob(ext):
                        if importer.import_module(file):
                            imported_count += 1
        
        if imported_count > 0:
            print(f"\n✓ {imported_count} Module importiert")
        else:
            print("\n⚠️  Keine Module gefunden oder importiert")

def main():
    parser = argparse.ArgumentParser(
        description='visiowings - VBA Editor für Visio mit VS Code Integration (Multi-Dokument-Support)',
        epilog='Beispiel: visiowings edit --file dokument.vsdx --force --bidirectional --debug'
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Verfügbare Befehle')
    
    # Edit command
    edit_parser = subparsers.add_parser(
        'edit', 
        help='VBA-Module bearbeiten mit Live-Sync (VS Code ↔ Visio)'
    )
    edit_parser.add_argument('--file', '-f', required=True, help='Visio-Datei (.vsdm, .vsdx, .vstm, .vstx)')
    edit_parser.add_argument('--output', '-o', help='Export-Verzeichnis (Standard: aktuelles Verzeichnis)')
    edit_parser.add_argument(
        '--force', 
        action='store_true', 
        help='Document-Module überschreiben (ThisDocument.cls)'
    )
    edit_parser.add_argument(
        '--bidirectional', 
        action='store_true', 
        help='Bidirektionaler Sync: Änderungen in Visio automatisch nach VS Code exportieren'
    )
    edit_parser.add_argument(
        '--debug',
        action='store_true',
        help='Debug-Modus: Ausführliche Log-Ausgaben'
    )
    
    # Export command
    export_parser = subparsers.add_parser('export', help='VBA-Module exportieren (einmalig)')
    export_parser.add_argument('--file', '-f', required=True, help='Visio-Datei (.vsdm, .vsdx, .vstm, .vstx)')
    export_parser.add_argument('--output', '-o', help='Export-Verzeichnis (Standard: aktuelles Verzeichnis)')
    export_parser.add_argument(
        '--debug',
        action='store_true',
        help='Debug-Modus: Ausführliche Log-Ausgaben'
    )
    
    # Import command
    import_parser = subparsers.add_parser('import', help='VBA-Module importieren (einmalig)')
    import_parser.add_argument('--file', '-f', required=True, help='Visio-Datei (.vsdm, .vsdx, .vstm, .vstx)')
    import_parser.add_argument('--input', '-i', help='Import-Verzeichnis (Standard: aktuelles Verzeichnis)')
    import_parser.add_argument(
        '--force', 
        action='store_true', 
        help='Document-Module überschreiben (ThisDocument.cls)'
    )
    import_parser.add_argument(
        '--debug',
        action='store_true',
        help='Debug-Modus: Ausführliche Log-Ausgaben'
    )
    
    args = parser.parse_args()
    
    if args.command == 'edit':
        cmd_edit(args)
    elif args.command == 'export':
        cmd_export(args)
    elif args.command == 'import':
        cmd_import(args)
    else:
        parser.print_help()

if __name__ == '__main__':
    main()
