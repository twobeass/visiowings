"""Event monitoring for Visio VBA debugging state.

Monitors VBA execution state changes and notifies the debug adapter
of break events, continue events, and execution completion.
"""

import asyncio
import logging
import threading
from enum import Enum
from typing import Callable, Optional

try:
    import pythoncom
    import win32com.client
except ImportError:
    raise ImportError("pywin32 is required. Install with: pip install pywin32")

logger = logging.getLogger(__name__)


class VBAExecutionMode(Enum):
    """VBA execution modes."""
    DESIGN = "design"
    RUN = "run"
    BREAK = "break"
    UNKNOWN = "unknown"


class VBAEventMonitor:
    """Monitor VBA execution state and raise events.
    
    Uses polling to detect state changes in VBA environment since
    COM event sinks for VBE are not reliably available.
    """

    def __init__(self, com_bridge, poll_interval: float = 0.5):
        """Initialize event monitor.
        
        Args:
            com_bridge: COMBridge instance for accessing VBA state
            poll_interval: Polling interval in seconds
        """
        self.com_bridge = com_bridge
        self.poll_interval = poll_interval
        self._running = False
        self._monitor_thread: Optional[threading.Thread] = None
        self._last_mode = VBAExecutionMode.UNKNOWN
        self._callbacks = {
            'on_break': [],
            'on_continue': [],
            'on_stopped': [],
            'on_error': [],
        }
        
    def register_callback(self, event: str, callback: Callable):
        """Register a callback for an event.
        
        Args:
            event: Event name ('on_break', 'on_continue', 'on_stopped', 'on_error')
            callback: Callback function to invoke
        """
        if event in self._callbacks:
            self._callbacks[event].append(callback)
        else:
            logger.warning(f"Unknown event type: {event}")
    
    def start(self):
        """Start monitoring VBA state."""
        if self._running:
            return
            
        self._running = True
        self._monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self._monitor_thread.start()
        logger.info("VBA event monitor started")
    
    def stop(self):
        """Stop monitoring VBA state."""
        self._running = False
        if self._monitor_thread:
            self._monitor_thread.join(timeout=2)
        logger.info("VBA event monitor stopped")
    
    def _monitor_loop(self):
        """Main monitoring loop."""
        pythoncom.CoInitialize()
        
        try:
            while self._running:
                try:
                    current_mode = self._check_execution_mode()
                    
                    # Detect state transitions
                    if current_mode != self._last_mode:
                        self._handle_state_change(self._last_mode, current_mode)
                        self._last_mode = current_mode
                    
                    # Sleep between polls
                    threading.Event().wait(self.poll_interval)
                    
                except Exception as e:
                    logger.error(f"Error in monitor loop: {e}", exc_info=True)
                    self._invoke_callbacks('on_error', str(e))
                    
        finally:
            pythoncom.CoUninitialize()
    
    def _check_execution_mode(self) -> VBAExecutionMode:
        """Check current VBA execution mode.
        
        Returns:
            Current execution mode
        """
        try:
            if not self.com_bridge.vbe:
                return VBAExecutionMode.UNKNOWN
            
            # Try to determine mode from VBE state
            # Mode property: 0=Design, 1=Run, 2=Break
            try:
                mode = self.com_bridge.vbe.Mode
                mode_map = {
                    0: VBAExecutionMode.DESIGN,
                    1: VBAExecutionMode.RUN,
                    2: VBAExecutionMode.BREAK,
                }
                return mode_map.get(mode, VBAExecutionMode.UNKNOWN)
            except AttributeError:
                # If Mode property not available, try alternative detection
                return self._detect_mode_alternative()
                
        except Exception as e:
            logger.debug(f"Could not determine execution mode: {e}")
            return VBAExecutionMode.UNKNOWN
    
    def _detect_mode_alternative(self) -> VBAExecutionMode:
        """Alternative mode detection when VBE.Mode is not available.
        
        Returns:
            Detected execution mode
        """
        try:
            # Check if code pane is accessible and what state it's in
            if hasattr(self.com_bridge.vbe, 'ActiveCodePane'):
                active_pane = self.com_bridge.vbe.ActiveCodePane
                if active_pane:
                    # In break mode, we can typically access more properties
                    try:
                        # Try to get the current line (only available in break)
                        _ = active_pane.TopLine
                        return VBAExecutionMode.BREAK
                    except:
                        pass
            
            # Default to design mode if can't determine
            return VBAExecutionMode.DESIGN
            
        except Exception:
            return VBAExecutionMode.UNKNOWN
    
    def _handle_state_change(self, old_mode: VBAExecutionMode, 
                            new_mode: VBAExecutionMode):
        """Handle execution mode state change.
        
        Args:
            old_mode: Previous execution mode
            new_mode: New execution mode
        """
        logger.debug(f"VBA state changed: {old_mode.value} -> {new_mode.value}")
        
        # Entering break mode (hit breakpoint or paused)
        if new_mode == VBAExecutionMode.BREAK and old_mode != VBAExecutionMode.BREAK:
            location = self._get_current_location()
            self._invoke_callbacks('on_break', location)
        
        # Leaving break mode (continue execution)
        elif old_mode == VBAExecutionMode.BREAK and new_mode == VBAExecutionMode.RUN:
            self._invoke_callbacks('on_continue')
        
        # Execution stopped (back to design from run)
        elif old_mode == VBAExecutionMode.RUN and new_mode == VBAExecutionMode.DESIGN:
            self._invoke_callbacks('on_stopped')
    
    def _get_current_location(self) -> dict:
        """Get current execution location when in break mode.
        
        Returns:
            Dictionary with module, procedure, and line information
        """
        location = {
            'module': None,
            'procedure': None,
            'line': None,
        }
        
        try:
            if not self.com_bridge.vbe:
                return location
            
            active_pane = self.com_bridge.vbe.ActiveCodePane
            if not active_pane:
                return location
            
            # Get module name
            code_module = active_pane.CodeModule
            if code_module and hasattr(code_module, 'Parent'):
                location['module'] = code_module.Parent.Name
            
            # Get current line
            try:
                location['line'] = active_pane.TopLine
            except:
                pass
            
            # Try to get procedure name
            if code_module and location['line']:
                try:
                    proc_name = code_module.ProcOfLine(location['line'], 0)
                    location['procedure'] = proc_name
                except:
                    pass
            
        except Exception as e:
            logger.debug(f"Could not get current location: {e}")
        
        return location
    
    def _invoke_callbacks(self, event: str, *args):
        """Invoke all callbacks for an event.
        
        Args:
            event: Event name
            args: Arguments to pass to callbacks
        """
        for callback in self._callbacks.get(event, []):
            try:
                callback(*args)
            except Exception as e:
                logger.error(f"Error in {event} callback: {e}", exc_info=True)
