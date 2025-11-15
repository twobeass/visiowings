"""Remote VBA debugging support for visiowings.

Provides Debug Adapter Protocol (DAP) integration for debugging
Visio VBA code directly from VS Code.
"""

from .debug_adapter import VisioDebugAdapter
from .debug_session import DebugSession
from .com_bridge import COMBridge
from .breakpoint_manager import BreakpointManager
from .event_monitor import VBAEventMonitor, VBAExecutionMode
from .variable_inspector import VariableInspector
from .callstack_inspector import CallStackInspector
from .error_handler import ErrorHandler, BreakpointCleanupManager

__all__ = [
    'VisioDebugAdapter',
    'DebugSession',
    'COMBridge',
    'BreakpointManager',
    'VBAEventMonitor',
    'VBAExecutionMode',
    'VariableInspector',
    'CallStackInspector',
    'ErrorHandler',
    'BreakpointCleanupManager',
]

__version__ = '0.1.0'
