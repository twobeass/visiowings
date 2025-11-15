"""COM bridge for Visio VBA debugging.

Provides thread-safe COM automation interface to Visio's VBA environment
for debugging operations.
"""

import asyncio
import logging
import threading
from queue import Queue
from typing import Any, Dict, List, Optional, Tuple

try:
    import win32com.client
    import pythoncom
except ImportError:
    raise ImportError("pywin32 is required for COM bridge. Install with: pip install pywin32")

logger = logging.getLogger(__name__)


class COMBridge:
    """Thread-safe bridge for Visio COM automation.
    
    Manages COM communication with Visio VBA environment in a dedicated thread
    to ensure thread safety and proper COM apartment initialization.
    """

    def __init__(self):
        self.visio_app = None
        self.vbe = None
        self.active_project = None
        self._com_thread = None
        self._request_queue = Queue()
        self._response_queue = Queue()
        self._running = False
        self._lock = threading.Lock()
        
    def start(self):
        """Start the COM bridge thread."""
        if self._running:
            return
            
        self._running = True
        self._com_thread = threading.Thread(target=self._com_worker, daemon=True)
        self._com_thread.start()
        logger.info("COM bridge started")
    
    def stop(self):
        """Stop the COM bridge thread."""
        self._running = False
        if self._com_thread:
            self._com_thread.join(timeout=5)
        logger.info("COM bridge stopped")
    
    def _com_worker(self):
        """COM worker thread that handles all COM operations."""
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        try:
            while self._running:
                try:
                    # Check for requests with timeout
                    if not self._request_queue.empty():
                        operation, args, kwargs = self._request_queue.get(timeout=0.1)
                        
                        try:
                            result = self._execute_operation(operation, args, kwargs)
                            self._response_queue.put(('success', result))
                        except Exception as e:
                            logger.error(f"COM operation failed: {e}", exc_info=True)
                            self._response_queue.put(('error', str(e)))
                            
                except Exception as e:
                    logger.error(f"COM worker error: {e}")
                    
        finally:
            # Cleanup COM
            if self.visio_app:
                try:
                    self.visio_app = None
                except:
                    pass
            pythoncom.CoUninitialize()
    
    def _execute_operation(self, operation: str, args: tuple, kwargs: dict) -> Any:
        """Execute a COM operation.
        
        Args:
            operation: Operation name
            args: Positional arguments
            kwargs: Keyword arguments
            
        Returns:
            Operation result
        """
        method = getattr(self, f"_op_{operation}", None)
        if not method:
            raise ValueError(f"Unknown operation: {operation}")
        return method(*args, **kwargs)
    
    async def execute(self, operation: str, *args, timeout: float = 5.0, **kwargs) -> Any:
        """Execute a COM operation asynchronously.
        
        Args:
            operation: Operation name
            args: Positional arguments
            timeout: Timeout in seconds
            kwargs: Keyword arguments
            
        Returns:
            Operation result
            
        Raises:
            TimeoutError: If operation times out
            Exception: If operation fails
        """
        self._request_queue.put((operation, args, kwargs))
        
        # Wait for response with timeout
        start_time = asyncio.get_event_loop().time()
        while asyncio.get_event_loop().time() - start_time < timeout:
            if not self._response_queue.empty():
                status, result = self._response_queue.get()
                if status == 'success':
                    return result
                else:
                    raise Exception(result)
            await asyncio.sleep(0.01)
        
        raise TimeoutError(f"COM operation '{operation}' timed out after {timeout}s")
    
    # COM Operations
    
    def _op_connect(self, visio_file: Optional[str] = None) -> bool:
        """Connect to Visio application.
        
        Args:
            visio_file: Optional Visio file to open
            
        Returns:
            True if successful
        """
        try:
            self.visio_app = win32com.client.Dispatch('Visio.Application')
            self.visio_app.Visible = True
            
            if visio_file:
                self.visio_app.Documents.Open(visio_file)
            
            # Access VBA environment
            self.vbe = self.visio_app.VBE
            if self.vbe.VBProjects.Count > 0:
                self.active_project = self.vbe.VBProjects(1)
            
            logger.info(f"Connected to Visio (VBA Projects: {self.vbe.VBProjects.Count})")
            return True
            
        except Exception as e:
            logger.error(f"Failed to connect to Visio: {e}")
            raise
    
    def _op_get_modules(self) -> List[Dict[str, str]]:
        """Get list of VBA modules.
        
        Returns:
            List of module information dictionaries
        """
        if not self.active_project:
            return []
        
        modules = []
        for component in self.active_project.VBComponents:
            modules.append({
                'name': component.Name,
                'type': component.Type,
                'code_lines': component.CodeModule.CountOfLines,
            })
        
        return modules
    
    def _op_get_code(self, module_name: str) -> str:
        """Get code from a VBA module.
        
        Args:
            module_name: Name of the module
            
        Returns:
            Module code
        """
        if not self.active_project:
            raise Exception("No active VBA project")
        
        component = self.active_project.VBComponents(module_name)
        code_module = component.CodeModule
        
        if code_module.CountOfLines == 0:
            return ""
        
        return code_module.Lines(1, code_module.CountOfLines)
    
    def _op_inject_breakpoint(self, module_name: str, line_number: int) -> Tuple[str, int]:
        """Inject a breakpoint at specified line.
        
        Args:
            module_name: Module name
            line_number: Line number (1-based)
            
        Returns:
            Tuple of (original_line, actual_line_number)
        """
        if not self.active_project:
            raise Exception("No active VBA project")
        
        component = self.active_project.VBComponents(module_name)
        code_module = component.CodeModule
        
        # Get original line
        original_line = code_module.Lines(line_number, 1)
        
        # Replace with Stop statement
        code_module.ReplaceLine(line_number, "Stop '" + original_line.strip())
        
        logger.debug(f"Injected breakpoint at {module_name}:{line_number}")
        return original_line, line_number
    
    def _op_remove_breakpoint(self, module_name: str, line_number: int, 
                             original_line: str) -> bool:
        """Remove a breakpoint and restore original code.
        
        Args:
            module_name: Module name
            line_number: Line number
            original_line: Original code line to restore
            
        Returns:
            True if successful
        """
        if not self.active_project:
            raise Exception("No active VBA project")
        
        try:
            component = self.active_project.VBComponents(module_name)
            code_module = component.CodeModule
            code_module.ReplaceLine(line_number, original_line)
            logger.debug(f"Removed breakpoint at {module_name}:{line_number}")
            return True
        except Exception as e:
            logger.error(f"Failed to remove breakpoint: {e}")
            return False
    
    def _op_get_debug_state(self) -> Dict[str, Any]:
        """Get current VBA debug state.
        
        Returns:
            Dictionary with debug state information
        """
        if not self.vbe:
            return {'mode': 'unknown'}
        
        try:
            # Try to access debugger state
            # Note: Actual implementation depends on Visio's COM API capabilities
            return {
                'mode': 'design',  # Could be 'design', 'run', 'break'
                'active': False,
            }
        except:
            return {'mode': 'unknown'}
    
    def _op_get_call_stack(self) -> List[Dict[str, Any]]:
        """Get current call stack.
        
        Returns:
            List of stack frame dictionaries
        """
        # This is a placeholder - actual implementation depends on
        # what Visio's COM API exposes for runtime inspection
        return []
    
    def _op_evaluate_expression(self, expression: str) -> Any:
        """Evaluate a VBA expression.
        
        Args:
            expression: VBA expression to evaluate
            
        Returns:
            Evaluation result
        """
        # This would use VBA's Immediate window or similar mechanism
        # Actual implementation depends on Visio's COM capabilities
        logger.warning("Expression evaluation not fully implemented yet")
        return None
    
    def _op_step_over(self) -> bool:
        """Execute step over command.
        
        Returns:
            True if successful
        """
        try:
            # Use VBA Commands if available, otherwise SendKeys
            # This is a simplified version
            import win32api
            import win32con
            
            # F8 key for step over in VBA
            win32api.keybd_event(win32con.VK_F8, 0, 0, 0)
            win32api.keybd_event(win32con.VK_F8, 0, win32con.KEYEVENTF_KEYUP, 0)
            return True
        except Exception as e:
            logger.error(f"Step over failed: {e}")
            return False
    
    def _op_step_in(self) -> bool:
        """Execute step in command.
        
        Returns:
            True if successful
        """
        try:
            import win32api
            import win32con
            
            # Shift+F8 for step in
            win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
            win32api.keybd_event(win32con.VK_F8, 0, 0, 0)
            win32api.keybd_event(win32con.VK_F8, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)
            return True
        except Exception as e:
            logger.error(f"Step in failed: {e}")
            return False
    
    def _op_step_out(self) -> bool:
        """Execute step out command.
        
        Returns:
            True if successful
        """
        try:
            import win32api
            import win32con
            
            # Ctrl+Shift+F8 for step out
            win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
            win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
            win32api.keybd_event(win32con.VK_F8, 0, 0, 0)
            win32api.keybd_event(win32con.VK_F8, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
            return True
        except Exception as e:
            logger.error(f"Step out failed: {e}")
            return False
    
    def _op_continue(self) -> bool:
        """Continue execution.
        
        Returns:
            True if successful
        """
        try:
            import win32api
            import win32con
            
            # F5 to continue
            win32api.keybd_event(win32con.VK_F5, 0, 0, 0)
            win32api.keybd_event(win32con.VK_F5, 0, win32con.KEYEVENTF_KEYUP, 0)
            return True
        except Exception as e:
            logger.error(f"Continue failed: {e}")
            return False
