"""CLI commands for VBA debugging.

Provides command-line interface for starting and managing
the debug adapter server.
"""

import argparse
import asyncio
import logging
import sys
from pathlib import Path

from .debug_adapter import VisioDebugAdapter

logger = logging.getLogger(__name__)


def setup_logging(verbose: bool = False):
    """Setup logging configuration.
    
    Args:
        verbose: Enable verbose logging
    """
    level = logging.DEBUG if verbose else logging.INFO
    
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler('visiowings-debug.log'),
        ]
    )


def cmd_start_server(args):
    """Start the debug adapter server.
    
    Args:
        args: Command arguments
    """
    setup_logging(args.verbose)
    
    logger.info(f"Starting debug adapter server on {args.host}:{args.port}")
    
    adapter = VisioDebugAdapter()
    
    try:
        asyncio.run(adapter.start_server(args.host, args.port))
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server error: {e}", exc_info=True)
        sys.exit(1)


def cmd_version(args):
    """Print version information.
    
    Args:
        args: Command arguments
    """
    from visiowings import __version__
    print(f"visiowings {__version__}")
    print("VBA Remote Debugging Support")


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description='Visio VBA Remote Debugging Tools',
        prog='visiowings-debug'
    )
    
    parser.add_argument(
        '--version',
        action='version',
        version='%(prog)s 0.1.0'
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Available commands')
    
    # Start server command
    server_parser = subparsers.add_parser(
        'start',
        help='Start debug adapter server'
    )
    server_parser.add_argument(
        '--host',
        default='127.0.0.1',
        help='Host address to bind to (default: 127.0.0.1)'
    )
    server_parser.add_argument(
        '--port',
        type=int,
        default=5678,
        help='Port number to listen on (default: 5678)'
    )
    server_parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Enable verbose logging'
    )
    server_parser.set_defaults(func=cmd_start_server)
    
    # Version command
    version_parser = subparsers.add_parser(
        'version',
        help='Show version information'
    )
    version_parser.set_defaults(func=cmd_version)
    
    # Parse and execute
    args = parser.parse_args()
    
    if hasattr(args, 'func'):
        args.func(args)
    else:
        parser.print_help()


if __name__ == '__main__':
    main()
