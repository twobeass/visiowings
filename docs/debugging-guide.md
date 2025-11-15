# VBA Remote Debugging Guide

This guide explains how to use visiowings' remote debugging feature to debug Visio VBA code directly from VS Code.

## Overview

The remote debugging system enables:

- **Full VS Code Integration**: Debug VBA code using VS Code's powerful debugging UI
- **Breakpoint Management**: Set, remove, and manage breakpoints visually
- **Step Execution**: Step over, into, and out of VBA code
- **Variable Inspection**: View local and global variables (limited support)
- **Call Stack**: Navigate the VBA call stack
- **Session Resilience**: Reconnect to debugging sessions after VS Code restart

## Requirements

### System Requirements

- **Operating System**: Windows only (COM automation required)
- **Visio**: Microsoft Visio (tested with 2016, 2019, 2021)
- **Python**: Python 3.8 or higher
- **VS Code**: Visual Studio Code with Python extension

### Python Dependencies

Install debug-specific dependencies:

```bash
pip install -r requirements-debug.txt
```

Or install manually:

```bash
pip install pywin32 asyncio
```

## Installation

### 1. Install visiowings with Debug Support

```bash
pip install visiowings[debug]
```

Or from source:

```bash
git clone https://github.com/twobeass/visiowings.git
cd visiowings
pip install -e .[debug]
```

### 2. Configure pywin32

After installing pywin32, run the post-install script:

```bash
python Scripts/pywin32_postinstall.py -install
```

### 3. Configure VS Code

Copy the `.vscode/launch.json` configuration to your workspace:

```json
{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Debug Visio VBA",
            "type": "visiowings",
            "request": "launch",
            "visioFile": "${workspaceFolder}/your-file.vsd",
            "stopOnEntry": false,
            "trace": true
        }
    ]
}
```

## Usage

### Starting the Debug Server

Start the debug adapter server:

```bash
visiowings-debug start
```

With custom host/port:

```bash
visiowings-debug start --host 127.0.0.1 --port 5678 --verbose
```

The server will listen for connections from VS Code.

### Debugging Workflow

#### 1. Launch Mode (Open Visio File)

1. Start the debug server: `visiowings-debug start`
2. Open your VBA source folder in VS Code
3. Set breakpoints in `.bas` files
4. Press F5 or select "Debug Visio VBA (Launch)" from Run menu
5. VS Code will connect to Visio and inject breakpoints
6. Run your VBA code in Visio
7. Execution pauses at breakpoints

#### 2. Attach Mode (Connect to Running Visio)

1. Open your Visio file manually
2. Start the debug server
3. In VS Code, select "Debug Visio VBA (Attach)"
4. Set breakpoints
5. Continue debugging as above

### Debugging Operations

#### Setting Breakpoints

- Click in the gutter next to line numbers in VS Code
- Breakpoints are injected as `Stop` statements in VBA
- Original code is preserved and restored on removal

#### Stepping Through Code

- **F10**: Step Over - Execute current line
- **F11**: Step Into - Enter function/sub calls
- **Shift+F11**: Step Out - Exit current function/sub
- **F5**: Continue - Resume execution

#### Inspecting Variables

- Hover over variables to see values (limited support)
- Use the Variables pane to view locals and globals
- Use Debug Console to evaluate expressions

#### Call Stack

- View the call stack in the Call Stack pane
- Click frames to navigate code

## Architecture

### Components

```
┌─────────────────┐
│   VS Code UI    │
└────────┬────────┘
         │ DAP Protocol
┌────────▼────────────┐
│  Debug Adapter      │  (Python asyncio)
│  VisioDebugAdapter  │
└────────┬────────────┘
         │
┌────────▼────────────┐
│  Debug Session      │  (Coordinator)
│  DebugSession       │
└────┬────────────┬───┘
     │            │
┌────▼────┐  ┌───▼──────────┐
│ COM     │  │ Breakpoint   │
│ Bridge  │  │ Manager      │
└────┬────┘  └──────────────┘
     │
┌────▼──────────────┐
│  Visio VBA COM    │
│  VBE.VBProjects   │
└───────────────────┘
```

### Thread Safety

- COM operations run in a dedicated thread with proper initialization
- Request/response queues mediate between async and COM contexts
- All COM calls have configurable timeouts (default: 5 seconds)

## Troubleshooting

### Common Issues

#### "Failed to connect to Visio"

**Cause**: Visio not running or COM automation disabled

**Solution**:
- Ensure Visio is installed and licensed
- Check Visio Trust Center settings (enable macros)
- Run VS Code as Administrator if needed

#### "Breakpoint not verified"

**Cause**: Module locked or code line not executable

**Solution**:
- Ensure VBA project is not password-protected
- Set breakpoints on executable lines (not comments/declarations)
- Check module is not read-only

#### "Step commands not working"

**Cause**: Visio VBA window not focused or SendKeys timing issues

**Solution**:
- Ensure Visio is in break mode
- Adjust timing delays in configuration
- Check for keyboard layout conflicts

#### "Variables not showing values"

**Cause**: Limited COM API access to runtime state

**Solution**:
- Use VBA's Immediate window for complex inspection
- Add debug output to VBA code
- Note: Full variable inspection is a known limitation

### Logging

Enable verbose logging:

```bash
visiowings-debug start --verbose
```

Logs are written to:
- Console output
- `visiowings-debug.log` in working directory

### Known Limitations

1. **Password-Protected Projects**: Cannot debug VBA projects with password protection
2. **Variable Inspection**: Limited runtime variable access via COM
3. **Conditional Breakpoints**: Not yet supported
4. **Watch Expressions**: Limited support
5. **Windows Only**: COM automation requires Windows
6. **Timing-Dependent**: SendKeys operations may fail under high load

## Advanced Configuration

### Custom Debug Server Settings

Create `visiowings-debug.json` in workspace root:

```json
{
    "host": "127.0.0.1",
    "port": 5678,
    "timeout": 10.0,
    "stepDelay": 0.1,
    "maxReconnectAttempts": 3
}
```

### VBA Project Setup

For best results:

1. **Enable Trust Access**: File → Options → Trust Center → Trust Center Settings → Macro Settings → "Trust access to the VBA project object model"

2. **Disable Password Protection**: Remove VBA project passwords during debugging

3. **Backup Code**: Breakpoint injection modifies code temporarily - always commit changes before debugging

### Security Considerations

- Breakpoint injection temporarily modifies VBA code
- Original code is restored on removal and session end
- No code modifications persist after debugging
- Temporary backups are not written to disk by default
- To enable encrypted backups, set `enableBackups: true` in config

## Examples

### Example 1: Debug a Simple Macro

**VBA Code (Module1.bas)**:
```vba
Sub TestMacro()
    Dim x As Integer
    x = 10
    x = x * 2
    MsgBox "Result: " & x
End Sub
```

**Steps**:
1. Export to `Module1.bas`
2. Open in VS Code
3. Set breakpoint on `x = x * 2`
4. Start debugging
5. Run `TestMacro` in Visio
6. Execution pauses at breakpoint

### Example 2: Debug Shape Event Handler

**VBA Code (ThisDocument.bas)**:
```vba
Private Sub Document_ShapeAdded(ByVal Shape As IVShape)
    Debug.Print "Shape added: " & Shape.Name
    ' Set breakpoint here
    Shape.Text = "New Shape"
End Sub
```

**Steps**:
1. Attach debugger to running Visio
2. Set breakpoint in event handler
3. Add shape in Visio
4. Debugger pauses in VS Code

## Contributing

Contributions to improve debugging support are welcome! Areas for enhancement:

- Enhanced variable inspection
- Conditional breakpoint support
- Better call stack introspection
- Watch expressions
- Hot reload support

See [CONTRIBUTING.md](../CONTRIBUTING.md) for guidelines.

## Support

For issues and questions:

- **GitHub Issues**: https://github.com/twobeass/visiowings/issues
- **Discussions**: https://github.com/twobeass/visiowings/discussions

## References

- [Debug Adapter Protocol](https://microsoft.github.io/debug-adapter-protocol/)
- [VS Code Debugging](https://code.visualstudio.com/docs/editor/debugging)
- [Visio VBA Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/visio)
- [pywin32 Documentation](https://github.com/mhammond/pywin32)
