import sys
from pathlib import Path

from .cli import cmd_edit, cmd_export, cmd_import


class InteractiveArgs:
    """Mock argparse arguments for interactive mode"""
    def __init__(self, **kwargs):
        for k, v in kwargs.items():
            setattr(self, k, v)

def prompt_bool(question, default=False):
    """Prompt for yes/no question"""
    suffix = " [Y/n]" if default else " [y/N]"
    ans = input(f"{question}{suffix}: ").strip().lower()
    if not ans:
        return default
    return ans in ('y', 'yes')

def prompt_path(question, default=None, must_exist=True):
    """Prompt for a file/directory path"""
    while True:
        default_str = f" [{default}]" if default else ""
        path_str = input(f"{question}{default_str}: ").strip()

        if not path_str and default:
            path_str = default

        if not path_str:
            print("⚠️  Path cannot be empty.")
            continue

        path = Path(path_str).resolve()
        if must_exist and not path.exists():
            print(f"⚠️  Path does not exist: {path}")
            continue

        return str(path)

def prompt_string(question, default=None):
    """Prompt for a string value"""
    default_str = f" [{default}]" if default else ""
    val = input(f"{question}{default_str}: ").strip()
    return val if val else default

def interactive_menu():
    print("\n🦋 Welcome to visiowings Interactive Mode")
    print("========================================")
    print("1. Edit Mode (Live Sync VS Code <-> Visio)")
    print("2. Export Modules (Visio -> Disk)")
    print("3. Import Modules (Disk -> Visio)")
    print("q. Quit")

    choice = input("\nSelect command: ").strip().lower()

    if choice in ('q', 'quit', 'exit'):
        sys.exit(0)

    visio_file = prompt_path("Path to Visio file (.vsdm, .vsdx)", must_exist=True)

    # Common arguments
    common_args = {
        'file': visio_file,
        'debug': prompt_bool("Enable debug mode?", default=False),
        'codepage': prompt_string("VBA Codepage (optional, e.g. cp1252)", default=None)
    }

    if choice == '1': # Edit
        output_dir = prompt_path("Export directory (for local files)", default=".", must_exist=False)
        bidirectional = prompt_bool("Enable bidirectional sync?", default=True)
        force = prompt_bool("Overwrite Document modules (ThisDocument)?", default=False)
        sync_delete = prompt_bool("Sync module deletions (delete in Visio if deleted locally)?", default=False)

        args = InteractiveArgs(
            command='edit',
            output=output_dir,
            bidirectional=bidirectional,
            force=force,
            sync_delete_modules=sync_delete,
            **common_args
        )
        print("\n🚀 Starting Edit Mode...\n")
        cmd_edit(args)

    elif choice == '2': # Export
        output_dir = prompt_path("Export directory", default=".", must_exist=False)

        args = InteractiveArgs(
            command='export',
            output=output_dir,
            **common_args
        )
        print("\n📤 Starting Export...\n")
        cmd_export(args)

    elif choice == '3': # Import
        input_dir = prompt_path("Import directory (source files)", default=".", must_exist=True)
        force = prompt_bool("Overwrite Document modules (ThisDocument)?", default=False)

        args = InteractiveArgs(
            command='import',
            input=input_dir,
            force=force,
            **common_args
        )
        print("\n📥 Starting Import...\n")
        cmd_import(args)

    else:
        print("❌ Invalid selection")
        sys.exit(1)
