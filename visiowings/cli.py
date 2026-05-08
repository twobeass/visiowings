"""Command Line Interface for visiowings with bidirectional sync support
Now supports multiple documents (drawings + stencils)
Extended: 'edit' command with --sync-delete-modules
"""
from __future__ import annotations

import argparse
import codecs
import logging
import os
import sys
from pathlib import Path

from . import __version__
from ._logging import setup_logging
from .exceptions import (
    InvalidVisioFileError,
    UnsupportedEncodingError,
    VisiowingsError,
)
from .file_watcher import VBAWatcher
from .vba_export import VisioVBAExporter
from .vba_import import VisioVBAImporter

logger = logging.getLogger("visiowings.cli")


# --------------------------------------------------------------------------- #
# Validation helpers
# --------------------------------------------------------------------------- #
def _validate_visio_file(path: Path) -> Path:
    """Resolve and validate a Visio file path, raising on bad input."""

    resolved = path.resolve()
    if resolved.suffix.lower() not in InvalidVisioFileError.SUPPORTED_SUFFIXES:
        raise InvalidVisioFileError(str(resolved))
    if not resolved.exists():
        # Mirror the historical error message to stay friendly. The CLI
        # prints `exc.message` so the user sees this verbatim.
        from .exceptions import DocumentNotFoundError

        raise DocumentNotFoundError(str(resolved))
    return resolved


def _validate_codepage(name: str | None) -> str | None:
    if not name:
        return None
    try:
        codecs.lookup(name)
    except LookupError as exc:
        raise UnsupportedEncodingError(name) from exc
    return name


def _validate_writable_dir(path: Path, *, label: str) -> Path:
    """Resolve a directory and ensure it exists and is writable."""

    resolved = path.resolve()
    resolved.mkdir(parents=True, exist_ok=True)
    if not os.access(resolved, os.W_OK):
        raise VisiowingsError(f"{label} is not writable: {resolved}")
    return resolved


def _validate_readable_dir(path: Path, *, label: str) -> Path:
    resolved = path.resolve()
    if not resolved.exists():
        raise VisiowingsError(f"{label} does not exist: {resolved}")
    if not resolved.is_dir():
        raise VisiowingsError(f"{label} is not a directory: {resolved}")
    if not os.access(resolved, os.R_OK):
        raise VisiowingsError(f"{label} is not readable: {resolved}")
    return resolved


# --------------------------------------------------------------------------- #
# Subcommands
# --------------------------------------------------------------------------- #
def cmd_edit(args):
    """Edit command: Export + Watch + Import with live sync"""
    visio_file = _validate_visio_file(Path(args.file))
    output_dir = _validate_writable_dir(Path(args.output or "."), label="--output")
    debug = getattr(args, "debug", False)
    sync_delete_modules = getattr(args, "sync_delete_modules", False)
    codepage = _validate_codepage(getattr(args, "codepage", None))
    use_rubberduck = getattr(args, "rubberduck", False)

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
    visio_file = _validate_visio_file(Path(args.file))
    output_dir = _validate_writable_dir(Path(args.output or "."), label="--output")
    debug = getattr(args, "debug", False)
    codepage = _validate_codepage(getattr(args, "codepage", None))
    use_rubberduck = getattr(args, "rubberduck", False)

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
    visio_file = _validate_visio_file(Path(args.file))
    input_dir = _validate_readable_dir(Path(args.input or "."), label="--input")
    debug = getattr(args, "debug", False)
    codepage = _validate_codepage(getattr(args, "codepage", None))
    use_rubberduck = getattr(args, "rubberduck", False)

    if use_rubberduck:
        print("🦆 Rubberduck integration enabled (folder annotations)")

    importer = VisioVBAImporter(str(visio_file), force_document=args.force, debug=debug, user_codepage=codepage, use_rubberduck=use_rubberduck)

    # Use new batch import method
    imported_count = importer.import_modules_from_dir(input_dir)

    if imported_count > 0:
        print(f"\n✓ {imported_count} modules imported")
    else:
        print("\n⚠️  No modules imported (or all skipped)")


# --------------------------------------------------------------------------- #
# Argument parser
# --------------------------------------------------------------------------- #
def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="visiowings",
        description="visiowings - VBA Editor for Visio with VS Code Integration (Multi-Document Support)",
        epilog="Example: visiowings edit --file document.vsdx --force --bidirectional --debug",
    )
    parser.add_argument(
        "--version", "-V",
        action="version",
        version=f"%(prog)s {__version__}",
    )

    subparsers = parser.add_subparsers(dest="command", help="Available commands")

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

    return parser


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #
def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)

    # If no arguments are passed, start interactive menu
    if not argv:
        from .interactive import interactive_menu
        interactive_menu()
        return 0

    parser = _build_parser()
    args = parser.parse_args(argv)

    setup_logging(debug=getattr(args, "debug", False))

    try:
        if args.command == 'edit':
            cmd_edit(args)
        elif args.command == 'export':
            cmd_export(args)
        elif args.command == 'import':
            cmd_import(args)
        else:
            parser.print_help()
            return 0
    except VisiowingsError as exc:
        # User-facing message only; full traceback only in --debug mode.
        sys.stderr.write(f"❌ {exc.message}\n")
        if getattr(args, "debug", False):
            logger.exception("Traceback (--debug):")
        return 1
    except KeyboardInterrupt:
        sys.stderr.write("\n⏹  Interrupted by user.\n")
        return 130

    return 0


if __name__ == '__main__':
    raise SystemExit(main())
