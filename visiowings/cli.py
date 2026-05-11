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
from ._update_check import schedule_async as schedule_update_check
from .config import (
    CONFIG_FILENAME,
    VisiowingsConfig,
    load_config,
    write_config,
)
from .exceptions import (
    InvalidVisioFileError,
    UnsupportedEncodingError,
    VisiowingsError,
)

# COM-touching modules (file_watcher, vba_export, vba_import,
# document_manager, visio_connection) eager-import pywin32 at module top
# level. We defer those imports until a subcommand actually needs them
# so the entry point — including `visiowings --help` and
# `visiowings init --non-interactive` — also runs on systems without
# pywin32 (Linux/macOS dev boxes, CI Linux runners).

logger = logging.getLogger("visiowings.cli")


def _apply_config_defaults(args: argparse.Namespace, cfg: VisiowingsConfig) -> None:
    """Fill in missing args from a loaded ``.visiowings.toml`` config."""

    for field_name in ("file", "output", "input", "codepage"):
        if getattr(args, field_name, None) is None:
            value = getattr(cfg, field_name, None)
            if value is not None:
                setattr(args, field_name, value)
                logger.debug("config-layer: %s <- .visiowings.toml (%r)", field_name, value)
    for field_name in ("bidirectional", "rubberduck", "sync_delete_modules", "force"):
        # Argparse store_true defaults to False, so we only upgrade False -> True.
        if not getattr(args, field_name, False) and getattr(cfg, field_name, False):
            setattr(args, field_name, True)
            logger.debug("config-layer: %s <- .visiowings.toml (True)", field_name)


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
    from .file_watcher import VBAWatcher
    from .vba_export import VisioVBAExporter
    from .vba_import import VisioVBAImporter

    visio_file = _validate_visio_file(Path(args.file))
    output_dir = _validate_writable_dir(Path(args.output or "."), label="--output")
    debug = getattr(args, "debug", False)
    sync_delete_modules = getattr(args, "sync_delete_modules", False)
    codepage = _validate_codepage(getattr(args, "codepage", None))
    use_rubberduck = getattr(args, "rubberduck", False)

    logger.debug("cmd_edit starting: file=%s output=%s", visio_file, output_dir)
    logger.debug(
        "bidirectional=%s sync_delete_modules=%s codepage=%s rubberduck=%s",
        getattr(args, "bidirectional", False),
        sync_delete_modules,
        codepage or "<auto>",
        use_rubberduck,
    )

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
        force_export_frx=getattr(args, "export_frx", False),
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
    importer = VisioVBAImporter(
        str(visio_file),
        force_document=args.force,
        debug=debug,
        silent_reconnect=True,
        user_codepage=codepage,
        use_rubberduck=use_rubberduck,
    )
    if not importer.connect_to_visio():
        return

    watcher = VBAWatcher(
        output_dir,
        importer,
        exporter=exporter,
        bidirectional=getattr(args, "bidirectional", False),
        debug=debug,
        sync_delete_modules=sync_delete_modules,
    )
    watcher.last_export_hashes = all_hashes  # Fix: Transfer initial export hash to watcher
    watcher.start()


def cmd_export(args):
    """Export command: Export VBA modules only"""
    from .vba_export import VisioVBAExporter

    visio_file = _validate_visio_file(Path(args.file))
    output_dir = _validate_writable_dir(Path(args.output or "."), label="--output")
    debug = getattr(args, "debug", False)
    codepage = _validate_codepage(getattr(args, "codepage", None))
    use_rubberduck = getattr(args, "rubberduck", False)

    logger.debug("cmd_export starting: file=%s output=%s", visio_file, output_dir)
    logger.debug("codepage=%s rubberduck=%s", codepage or "<auto>", use_rubberduck)

    if use_rubberduck:
        print("🦆 Rubberduck integration enabled (folder annotations)")

    exporter = VisioVBAExporter(
        str(visio_file),
        debug=debug,
        user_codepage=codepage,
        use_rubberduck=use_rubberduck,
        force_export_frx=getattr(args, "export_frx", False),
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


_DEFAULT_INIT_FILE = "your-document.vsdm"
_DEFAULT_INIT_OUTPUT = "./vba"
_DEFAULT_INIT_CODEPAGE = "cp1252"


def cmd_init(args):
    """Generate a ``.visiowings.toml`` config in the current directory.

    Two modes:

    - **Interactive (default):** small wizard. Lists currently open Visio
      documents (via COM) and prompts for output dir, codepage, sync flags.
    - **Non-interactive** (``--non-interactive`` / ``-y``): skip all prompts
      and write a sensible default config. Required values can be passed in
      via ``--file``, ``--output``, ``--codepage`` flags. This is the path
      the UAT runner uses.

    Idempotency: if ``.visiowings.toml`` already exists the command refuses
    to overwrite. Pass ``--force`` to overwrite, or in interactive mode
    answer ``y`` to the confirm prompt.
    """

    config_path = Path.cwd() / CONFIG_FILENAME
    non_interactive = getattr(args, "non_interactive", False)
    force = getattr(args, "force", False)

    logger.debug(
        "cmd_init starting: cwd=%s non_interactive=%s force=%s exists=%s",
        Path.cwd(),
        non_interactive,
        force,
        config_path.exists(),
    )

    if config_path.exists() and not force:
        if non_interactive:
            raise VisiowingsError(
                f"{config_path} already exists. "
                f"Pass --force to overwrite (non-interactive mode does not prompt)."
            )
        answer = input(f"{config_path.name} already exists. Overwrite? [y/N]: ").strip().lower()
        if answer not in ("y", "yes"):
            print("Aborted; nothing written.")
            return
        force = True

    if non_interactive:
        cfg = VisiowingsConfig(
            file=getattr(args, "file", None) or _DEFAULT_INIT_FILE,
            output=getattr(args, "output", None) or _DEFAULT_INIT_OUTPUT,
            codepage=getattr(args, "codepage", None) or _DEFAULT_INIT_CODEPAGE,
            bidirectional=False,
            rubberduck=False,
        )
        target = write_config(cfg)
        print(f"✓ Wrote {target} (non-interactive defaults — edit before use)")
        return

    print(f"\n🦉 visiowings init — writing {config_path.name}\n")

    docs = _discover_open_documents()

    if docs:
        print("Open Visio documents:")
        for i, full_path in enumerate(docs, start=1):
            print(f"  {i}. {full_path}")
        print(f"  {len(docs) + 1}. Enter a path manually")
        while True:
            choice = input(f"\nSelect document [1-{len(docs) + 1}]: ").strip()
            if choice.isdigit():
                idx = int(choice)
                if 1 <= idx <= len(docs):
                    main_file = docs[idx - 1]
                    break
                if idx == len(docs) + 1:
                    main_file = input("Path to Visio file: ").strip()
                    break
            print("⚠️  Invalid selection, try again.")
    else:
        print("(No documents detected via COM. You can still enter a path manually.)\n")
        main_file = input("Path to Visio file: ").strip()

    output_dir = (
        input(f"Output directory for VBA files [{_DEFAULT_INIT_OUTPUT}]: ").strip()
        or _DEFAULT_INIT_OUTPUT
    )
    bidir = input("Enable bidirectional sync (y/N)? ").strip().lower() == "y"
    rubberduck = input("Use Rubberduck @Folder annotations (y/N)? ").strip().lower() == "y"
    codepage = input(f"Codepage [{_DEFAULT_INIT_CODEPAGE}, blank = auto-detect]: ").strip() or None

    cfg = VisiowingsConfig(
        file=main_file,
        output=output_dir,
        codepage=codepage,
        bidirectional=bidir,
        rubberduck=rubberduck,
    )
    target = write_config(cfg)
    print(f"\n✓ Wrote {target}")
    print("  Run `visiowings edit` (no arguments) to start syncing with these defaults.")


def _discover_open_documents() -> list[str]:
    """Best-effort listing of open Visio documents — empty list on non-Windows."""

    try:
        import pythoncom
        import win32com.client

        try:
            pythoncom.CoInitialize()
        except pythoncom.com_error:
            # COM is already initialized for this thread (RPC_E_CHANGED_MODE
            # or similar) — that's fine for a read-only enumeration.
            pass

        try:
            visio = win32com.client.GetActiveObject("Visio.Application")
        except Exception:
            return []
        return [doc.FullName for doc in visio.Documents]
    except Exception as e:
        logger.debug("Could not enumerate Visio documents: %s", e)
        return []


def cmd_import(args):
    """Import command: Import VBA modules only"""
    from .vba_import import VisioVBAImporter

    visio_file = _validate_visio_file(Path(args.file))
    input_dir = _validate_readable_dir(Path(args.input or "."), label="--input")
    debug = getattr(args, "debug", False)
    codepage = _validate_codepage(getattr(args, "codepage", None))
    use_rubberduck = getattr(args, "rubberduck", False)

    logger.debug("cmd_import starting: file=%s input=%s", visio_file, input_dir)
    logger.debug(
        "force=%s codepage=%s rubberduck=%s",
        getattr(args, "force", False),
        codepage or "<auto>",
        use_rubberduck,
    )

    if use_rubberduck:
        print("🦆 Rubberduck integration enabled (folder annotations)")

    importer = VisioVBAImporter(
        str(visio_file),
        force_document=args.force,
        debug=debug,
        user_codepage=codepage,
        use_rubberduck=use_rubberduck,
    )

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
        epilog=(
            "Example: visiowings edit --file document.vsdx --force --bidirectional --debug\n\n"
            "Config layering: command-line flags override values in "
            "`.visiowings.toml`, which override built-in defaults. Run "
            "`visiowings init` to scaffold a config in the current directory."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--version",
        "-V",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    parser.add_argument(
        "--no-update-check",
        action="store_true",
        help="Skip the daily PyPI update check (also disabled by VISIOWINGS_NO_UPDATE_CHECK=1)",
    )

    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # init command
    init_parser = subparsers.add_parser(
        "init",
        help="Create a .visiowings.toml config in the current directory",
    )
    init_parser.add_argument(
        "--force",
        action="store_true",
        help="Overwrite an existing .visiowings.toml",
    )
    init_parser.add_argument(
        "--non-interactive",
        "-y",
        action="store_true",
        help=(
            "Skip prompts and write a default .visiowings.toml. Useful for CI and the UAT runner."
        ),
    )
    init_parser.add_argument(
        "--file",
        help="Visio file path to record in the config (non-interactive mode)",
    )
    init_parser.add_argument(
        "--output",
        help="Output directory to record in the config (default: ./vba)",
    )
    init_parser.add_argument(
        "--codepage",
        help="Codepage to record in the config (default: cp1252)",
    )
    init_parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable structured [DEBUG] logging on stderr and show tracebacks on failure",
    )

    # Edit command
    edit_parser = subparsers.add_parser(
        "edit", help="Edit VBA modules with live sync (VS Code <-> Visio)"
    )
    edit_parser.add_argument(
        "--file",
        "-f",
        help="Visio file (.vsdm, .vsdx, .vstm, .vstx) - falls back to .visiowings.toml",
    )
    edit_parser.add_argument("--output", "-o", help="Export directory (default: current directory)")
    edit_parser.add_argument(
        "--force", action="store_true", help="Overwrite document modules (ThisDocument.cls)"
    )
    edit_parser.add_argument(
        "--bidirectional",
        action="store_true",
        help="Bidirectional sync: Automatically export changes from Visio to VS Code",
    )
    edit_parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable structured [DEBUG] logging on stderr and show tracebacks on failure",
    )
    edit_parser.add_argument(
        "--sync-delete-modules",
        action="store_true",
        help="Automatically delete Visio modules when local .bas/.cls/.frm files are deleted",
    )
    edit_parser.add_argument(
        "--codepage",
        "--cp",
        help="VBA file codepage (e.g., cp1252=Western, cp1251=Cyrillic, cp1250=Central EU, cp936=Chinese). Default: auto-detect from document",
    )
    edit_parser.add_argument(
        "--rubberduck",
        "--rd",
        action="store_true",
        help="Use Rubberduck @Folder annotations for directory structure",
    )
    edit_parser.add_argument(
        "--export-frx",
        action="store_true",
        help="Force export of .frx files even if .frm code has not changed",
    )

    # Export command
    export_parser = subparsers.add_parser("export", help="Export VBA modules (one-time)")
    export_parser.add_argument(
        "--file",
        "-f",
        help="Visio file (.vsdm, .vsdx, .vstm, .vstx) - falls back to .visiowings.toml",
    )
    export_parser.add_argument(
        "--output", "-o", help="Export directory (default: current directory)"
    )
    export_parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable structured [DEBUG] logging on stderr and show tracebacks on failure",
    )
    export_parser.add_argument(
        "--codepage",
        "--cp",
        help="VBA file codepage (e.g., cp1252=Western, cp1251=Cyrillic, cp1250=Central EU, cp936=Chinese). Default: auto-detect from document",
    )
    export_parser.add_argument(
        "--rubberduck",
        "--rd",
        action="store_true",
        help="Use Rubberduck @Folder annotations for directory structure",
    )
    export_parser.add_argument(
        "--export-frx",
        action="store_true",
        help="Force export of .frx files even if .frm code has not changed",
    )

    # Import command
    import_parser = subparsers.add_parser("import", help="Import VBA modules (one-time)")
    import_parser.add_argument(
        "--file",
        "-f",
        help="Visio file (.vsdm, .vsdx, .vstm, .vstx) - falls back to .visiowings.toml",
    )
    import_parser.add_argument(
        "--input", "-i", help="Import directory (default: current directory)"
    )
    import_parser.add_argument(
        "--force", action="store_true", help="Overwrite document modules (ThisDocument.cls)"
    )
    import_parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable structured [DEBUG] logging on stderr and show tracebacks on failure",
    )
    import_parser.add_argument(
        "--codepage",
        "--cp",
        help="VBA file codepage (e.g., cp1252=Western, cp1251=Cyrillic, cp1250=Central EU, cp936=Chinese). Default: auto-detect from document",
    )
    import_parser.add_argument(
        "--rubberduck",
        "--rd",
        action="store_true",
        help="Use Rubberduck @Folder annotations for directory structure",
    )

    return parser


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #
def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)

    # If no arguments are passed, start interactive menu. We inject the
    # command callables instead of letting `interactive` import them, so
    # the two modules don't form an import cycle.
    if not argv:
        from .interactive import interactive_menu

        interactive_menu(cmd_edit, cmd_export, cmd_import)
        return 0

    parser = _build_parser()
    args = parser.parse_args(argv)

    setup_logging(debug=getattr(args, "debug", False))

    if not getattr(args, "no_update_check", False):
        schedule_update_check()

    # init does not consume the project config (it creates one), so we skip
    # auto-loading there.
    if args.command not in (None, "init"):
        try:
            cfg = load_config()
        except Exception as e:
            logger.warning("Could not read .visiowings.toml: %s", e)
            cfg = VisiowingsConfig()
        _apply_config_defaults(args, cfg)

        # File is mandatory after merge with config
        if getattr(args, "file", None) is None and args.command in ("edit", "export", "import"):
            parser.error("--file is required (or set `file = ...` in .visiowings.toml)")

    try:
        if args.command == "init":
            cmd_init(args)
        elif args.command == "edit":
            cmd_edit(args)
        elif args.command == "export":
            cmd_export(args)
        elif args.command == "import":
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


if __name__ == "__main__":
    raise SystemExit(main())
