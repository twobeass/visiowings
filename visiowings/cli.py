"""Command Line Interface for visiowings with bidirectional sync support
Now supports multiple documents (drawings + stencils)
Extended: 'edit' command with --sync-delete-modules
"""
import argparse
from pathlib import Path

from .file_watcher import VBAWatcher
from .vba_export import VisioVBAExporter
from .vba_import import VisioVBAImporter


def cmd_edit(args):
    """Edit command: Export + Watch + Import with live sync"""
    visio_file = Path(args.file).resolve()
    output_dir = Path(args.output or '.').resolve()
    debug = getattr(args, 'debug', False)
    sync_delete_modules = getattr(args, 'sync_delete_modules', False)
    codepage = getattr(args, 'codepage', None)
    use_rubberduck = getattr(args, 'rubberduck', False)

    if not visio_file.exists():
        print(f"❌ File not found: {visio_file}")
        return

    print(f"📂 Visio file: {visio_file}")
    print(f"📁 Export directory: {output_dir}")
    if debug:
        print("[DEBUG] Debug mode enabled")
    if codepage:
        print(f"📝 Codepage: {codepage}")
    if use_rubberduck:
        print("🦆 Rubberduck integration enabled (folder annotations)")

    print("\n=== Exporting VBA Modules ===")
    exporter = VisioVBAExporter(
        str(visio_file),
        debug=debug,
        user_codepage=codepage,
        use_rubberduck=use_rubberduck,
        force_export_frx=getattr(args, 'export_frx', False)
    )
    if not exporter.connect_to_visio():
        return

    # Export returns dict format for multi-document support
    all_exported, all_hashes = exporter.export_modules(output_dir)

    if not all_exported:
        print("❌ No modules exported")
        return

    # Count total exported files
    total_files = sum(len(files) for files in all_exported.values())
    total_docs = len(all_exported)

    if total_docs > 1:
        print(f"\n✓ {total_files} modules exported from {total_docs} documents")
    else:
        print(f"\n✓ {total_files} modules exported")

    if debug:
        for doc_folder, doc_hash in all_hashes.items():
            print(f"[DEBUG] {doc_folder}: Hash {doc_hash[:8]}...")

    print("\n=== Starting Live Synchronization ===")
    importer = VisioVBAImporter(str(visio_file), force_document=args.force, debug=debug, silent_reconnect=True, user_codepage=codepage, use_rubberduck=use_rubberduck)
    if not importer.connect_to_visio():
        return

    watcher = VBAWatcher(
        output_dir,
        importer,
        exporter=exporter,
        bidirectional=getattr(args, 'bidirectional', False),
        debug=debug,
        sync_delete_modules=sync_delete_modules
    )
    watcher.last_export_hashes = all_hashes  # Fix: Transfer initial export hash to watcher
    watcher.start()

def cmd_export(args):
    """Export command: Export VBA modules only"""
    visio_file = Path(args.file).resolve()
    output_dir = Path(args.output or '.').resolve()
    debug = getattr(args, 'debug', False)
    codepage = getattr(args, 'codepage', None)
    use_rubberduck = getattr(args, 'rubberduck', False)

    if use_rubberduck:
        print("🦆 Rubberduck integration enabled (folder annotations)")

    exporter = VisioVBAExporter(
        str(visio_file),
        debug=debug,
        user_codepage=codepage,
        use_rubberduck=use_rubberduck,
        force_export_frx=getattr(args, 'export_frx', False)
    )
    if exporter.connect_to_visio():
        all_exported, all_hashes = exporter.export_modules(output_dir)

        if all_exported:
            total_files = sum(len(files) for files in all_exported.values())
            total_docs = len(all_exported)

            if total_docs > 1:
                print(f"\n✓ {total_files} modules exported from {total_docs} documents")
            else:
                print(f"\n✓ {total_files} modules exported")

            if debug:
                for doc_folder, doc_hash in all_hashes.items():
                    print(f"[DEBUG] {doc_folder}: Hash {doc_hash[:8]}...")

def cmd_import(args):
    """Import command: Import VBA modules only"""
    visio_file = Path(args.file).resolve()
    input_dir = Path(args.input or '.').resolve()
    debug = getattr(args, 'debug', False)
    codepage = getattr(args, 'codepage', None)
    use_rubberduck = getattr(args, 'rubberduck', False)

    if use_rubberduck:
        print("🦆 Rubberduck integration enabled (folder annotations)")

    importer = VisioVBAImporter(str(visio_file), force_document=args.force, debug=debug, user_codepage=codepage, use_rubberduck=use_rubberduck)

    # Use new batch import method
    imported_count = importer.import_modules_from_dir(input_dir)

    if imported_count > 0:
        print(f"\n✓ {imported_count} modules imported")
    else:
        print("\n⚠️  No modules imported (or all skipped)")

def main():
    parser = argparse.ArgumentParser(
        description='visiowings - VBA Editor for Visio with VS Code Integration (Multi-Document Support)',
        epilog='Example: visiowings edit --file document.vsdx --force --bidirectional --debug'
    )

    subparsers = parser.add_subparsers(dest='command', help='Available commands')

    # Edit command
    edit_parser = subparsers.add_parser(
        'edit',
        help='Edit VBA modules with live sync (VS Code <-> Visio)'
    )
    edit_parser.add_argument('--file', '-f', required=True, help='Visio file (.vsdm, .vsdx, .vstm, .vstx)')
    edit_parser.add_argument('--output', '-o', help='Export directory (default: current directory)')
    edit_parser.add_argument(
        '--force',
        action='store_true',
        help='Overwrite document modules (ThisDocument.cls)'
    )
    edit_parser.add_argument(
        '--bidirectional',
        action='store_true',
        help='Bidirectional sync: Automatically export changes from Visio to VS Code'
    )
    edit_parser.add_argument(
        '--debug',
        action='store_true',
        help='Debug mode: Verbose logging output'
    )
    edit_parser.add_argument(
        '--sync-delete-modules',
        action='store_true',
        help='Automatically delete Visio modules when local .bas/.cls/.frm files are deleted'
    )
    edit_parser.add_argument(
        '--codepage', '--cp',
        help='VBA file codepage (e.g., cp1252=Western, cp1251=Cyrillic, cp1250=Central EU, cp936=Chinese). Default: auto-detect from document'
    )
    edit_parser.add_argument(
        '--rubberduck', '--rd',
        action='store_true',
        help='Use Rubberduck @Folder annotations for directory structure'
    )
    edit_parser.add_argument(
        '--export-frx',
        action='store_true',
        help='Force export of .frx files even if .frm code has not changed'
    )

    # Export command
    export_parser = subparsers.add_parser('export', help='Export VBA modules (one-time)')
    export_parser.add_argument('--file', '-f', required=True, help='Visio file (.vsdm, .vsdx, .vstm, .vstx)')
    export_parser.add_argument('--output', '-o', help='Export directory (default: current directory)')
    export_parser.add_argument(
        '--debug',
        action='store_true',
        help='Debug mode: Verbose logging output'
    )
    export_parser.add_argument(
        '--codepage', '--cp',
        help='VBA file codepage (e.g., cp1252=Western, cp1251=Cyrillic, cp1250=Central EU, cp936=Chinese). Default: auto-detect from document'
    )
    export_parser.add_argument(
        '--rubberduck', '--rd',
        action='store_true',
        help='Use Rubberduck @Folder annotations for directory structure'
    )
    export_parser.add_argument(
        '--export-frx',
        action='store_true',
        help='Force export of .frx files even if .frm code has not changed'
    )

    # Import command
    import_parser = subparsers.add_parser('import', help='Import VBA modules (one-time)')
    import_parser.add_argument('--file', '-f', required=True, help='Visio file (.vsdm, .vsdx, .vstm, .vstx)')
    import_parser.add_argument('--input', '-i', help='Import directory (default: current directory)')
    import_parser.add_argument(
        '--force',
        action='store_true',
        help='Overwrite document modules (ThisDocument.cls)'
    )
    import_parser.add_argument(
        '--debug',
        action='store_true',
        help='Debug mode: Verbose logging output'
    )
    import_parser.add_argument(
        '--codepage', '--cp',
        help='VBA file codepage (e.g., cp1252=Western, cp1251=Cyrillic, cp1250=Central EU, cp936=Chinese). Default: auto-detect from document'
    )
    import_parser.add_argument(
        '--rubberduck', '--rd',
        action='store_true',
        help='Use Rubberduck @Folder annotations for directory structure'
    )

    # If no arguments are passed, start interactive menu
    import sys
    if len(sys.argv) == 1:
        from .interactive import interactive_menu
        interactive_menu()
        return

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
