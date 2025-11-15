# visiowings Remote VBA Debugging

Debug Visio VBA code directly from VS Code with full debugging capabilities.

## Quick Start

### Install

```bash
pip install visiowings[debug]
```

### Start Debug Server

```bash
visiowings-debug start
```

### Configure VS Code

Create `.vscode/launch.json`:

```json
{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Debug Visio VBA",
            "type": "visiowings",
            "request": "launch",
            "visioFile": "${workspaceFolder}/your-file.vsd"
        }
    ]
}
```

### Debug

1. Set breakpoints in `.bas` files
2. Press F5
3. Run VBA code in Visio
4. Debug in VS Code!

## Features

- ✅ Breakpoints
- ✅ Step over/in/out
- ✅ Call stack
- ✅ Continue/Pause
- ⚠️ Variable inspection (limited)
- ⚠️ Expression evaluation (limited)

## Documentation

- [Complete Guide](../../docs/debugging-guide.md)
- [Examples](../../examples/debug/)
- [Architecture](../../feature-debug.md)

## Requirements

- Windows only
- Python 3.8+
- Visio 2016+
- pywin32

## Components

- **debug_adapter.py**: DAP server
- **com_bridge.py**: COM automation
- **breakpoint_manager.py**: Breakpoint handling
- **debug_session.py**: Session coordination
- **cli.py**: Command-line interface

## Support

Issues: https://github.com/twobeass/visiowings/issues
