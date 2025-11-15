# VBA Debugging Examples

This directory contains example VBA modules for testing visiowings remote debugging capabilities.

## Files

- **Module1.bas**: Example VBA module with various debugging scenarios

## Setup

### 1. Create Test Visio File

Create a new Visio document:

1. Open Microsoft Visio
2. Create a new blank drawing
3. Add a few shapes (rectangles, circles, etc.)
4. Save as `test.vsd` or `test.vsdx`

### 2. Import VBA Module

**Option A: Manual Import**
1. In Visio, press Alt+F11 to open VBA Editor
2. File → Import File
3. Select `Module1.bas`

**Option B: Using visiowings**
```bash
visiowings import test.vsd examples/debug/Module1.bas
```

### 3. Enable VBA Trust Settings

1. In Visio: File → Options → Trust Center → Trust Center Settings
2. Macro Settings → Enable all macros (for debugging only)
3. Check "Trust access to the VBA project object model"

## Debugging Scenarios

Each function in `Module1.bas` demonstrates different debugging features:

### Basic Debugging

**Function: `CalculateSum`**
- Simple function with local variables
- Test basic breakpoints
- Inspect input parameters and return value

**How to test:**
1. Set breakpoint on `result = a + b`
2. Run `CalculateSum(5, 10)` in Immediate window
3. Inspect values of `a`, `b`, and `result`

### Loop Debugging

**Function: `ProcessShapes`**
- Iterates over shapes in active page
- Test breakpoints inside loops
- Inspect collection items

**How to test:**
1. Set breakpoint inside the `For Each` loop
2. Run `ProcessShapes` from VBA or Visio UI
3. Step through each iteration
4. Inspect `shp` and `counter` variables

### Conditional Debugging

**Function: `GetShapeCategory`**
- Multiple conditional branches
- Test stepping through different paths

**How to test:**
1. Set breakpoints in each `If/ElseIf/Else` branch
2. Test with different shape types
3. Observe which branch executes

### Error Handling

**Function: `SafeShapeOperation`**
- Demonstrates error handling with `On Error GoTo`
- Test debugging with exceptions

**How to test:**
1. Set breakpoint in error handler
2. Run function with missing cell
3. Observe error capture and handling

### Call Stack

**Functions: `TestCallStack`, `Level1/2/3Function`**
- Nested function calls
- Test call stack navigation
- Step into/out of functions

**How to test:**
1. Set breakpoint in `Level3Function`
2. Run `TestCallStack`
3. View call stack in VS Code
4. Navigate between stack frames

### Variable Inspection

**Function: `TestVariableInspection`**
- Multiple variable types
- Test variable inspection in Variables pane

**How to test:**
1. Set breakpoint after variable assignments
2. Run function
3. Inspect Variables pane in VS Code
4. Hover over variables in editor

## Debugging Workflow Example

### Complete Walkthrough

1. **Start Debug Server**
   ```bash
   visiowings-debug start --verbose
   ```

2. **Configure VS Code**
   - Open workspace containing exported VBA files
   - Create `.vscode/launch.json` (see main docs)

3. **Set Breakpoints**
   - Open `Module1.bas` in VS Code
   - Click gutter next to line numbers to set breakpoints
   - Try setting multiple breakpoints in `ProcessShapes`

4. **Launch Debugger**
   - Press F5 or click "Run and Debug"
   - Select "Debug Visio VBA (Launch)"
   - Wait for connection

5. **Run Code**
   - In Visio, press Alt+F11 (VBA Editor)
   - Press F5 or run a macro
   - Execution pauses at first breakpoint

6. **Debug Operations**
   - **Continue (F5)**: Run to next breakpoint
   - **Step Over (F10)**: Execute current line
   - **Step Into (F11)**: Enter function calls
   - **Step Out (Shift+F11)**: Exit current function

7. **Inspect State**
   - View Variables pane for locals/globals
   - View Call Stack pane for execution context
   - Hover over variables in editor
   - Use Debug Console for expressions

8. **End Session**
   - Click Stop button or press Shift+F5
   - Breakpoints are automatically removed

## Tips

### Best Practices

- **Save Before Debugging**: Always save Visio file before starting debug session
- **One Session at a Time**: Close other Visio instances to avoid conflicts
- **Test Simple First**: Start with `CalculateSum` before complex scenarios
- **Check Logs**: Review `visiowings-debug.log` for troubleshooting

### Common Issues

**Breakpoint Not Hit**
- Ensure code is actually executed (not dead code)
- Check that module is imported correctly
- Verify breakpoint is on executable line (not comment)

**Variables Not Showing**
- This is a known limitation
- Use Debug.Print statements as workaround
- Check Immediate window in VBA Editor

**Step Commands Slow**
- SendKeys operations have timing delays
- Ensure Visio window is not minimized
- Close unnecessary applications

## Advanced Scenarios

### Debugging Event Handlers

Create event handler in `ThisDocument`:

```vba
Private Sub Document_ShapeAdded(ByVal Shape As IVShape)
    ' Set breakpoint here
    Debug.Print "Shape added: " & Shape.Name
End Sub
```

Test by adding shapes in Visio while debugging.

### Debugging Forms

For UserForms:
1. Export form with `visiowings export`
2. Set breakpoints in form code
3. Show form to trigger breakpoints

### Long-Running Operations

For operations that take time:
- Use pause command to break during execution
- Set breakpoints at key milestones
- Monitor progress with Debug.Print

## Resources

- [Main Debugging Guide](../../docs/debugging-guide.md)
- [DAP Specification](https://microsoft.github.io/debug-adapter-protocol/)
- [Visio VBA Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/visio)
