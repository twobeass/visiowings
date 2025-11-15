"""Debug session coordinator for Visio VBA debugging.

Manages the overall debugging session, coordinating between the debug adapter,
COM bridge, and breakpoint manager.
"""

import asyncio
import logging
from typing import Any, Dict, List, Optional

from .com_bridge import COMBridge
from .breakpoint_manager import BreakpointManager

logger = logging.getLogger(__name__)


class DebugSession:
    """Coordinates a VBA debugging session.
    
    Manages:
    - Connection to Visio VBA environment
    - Breakpoint operations
    - Execution control (step, continue, pause)
    - Variable and stack inspection
    - Session lifecycle and reconnection
    """
    
    def __init__(self, session_id: str, visio_file: Optional[str] = None):
        """Initialize debug session.
        
        Args:
            session_id: Unique session identifier
            visio_file: Optional path to Visio file
        """
        self.session_id = session_id
        self.visio_file = visio_file
        self.com_bridge = COMBridge()
        self.breakpoint_manager = BreakpointManager(self.com_bridge)
        self.is_running = False
        self.is_paused = False
        
    async def start(self):
        """Start the debug session."""
        logger.info(f"Starting debug session {self.session_id}")
        
        # Start COM bridge
        self.com_bridge.start()
        
        # Connect to Visio
        await self.com_bridge.execute('connect', self.visio_file)
        
        self.is_running = True
        logger.info(f"Debug session {self.session_id} started")
    
    async def attach(self):
        """Attach to existing Visio instance."""
        logger.info(f"Attaching debug session {self.session_id}")
        
        self.com_bridge.start()
        await self.com_bridge.execute('connect')
        
        self.is_running = True
        logger.info(f"Debug session {self.session_id} attached")
    
    async def reconnect(self):
        """Reconnect to existing session."""
        logger.info(f"Reconnecting debug session {self.session_id}")
        
        if not self.is_running:
            await self.attach()
        
        logger.info(f"Debug session {self.session_id} reconnected")
    
    async def disconnect(self):
        """Disconnect and cleanup session."""
        logger.info(f"Disconnecting debug session {self.session_id}")
        
        # Clear all breakpoints
        await self.breakpoint_manager.clear_all()
        
        # Stop COM bridge
        self.com_bridge.stop()
        
        self.is_running = False
        logger.info(f"Debug session {self.session_id} disconnected")
    
    async def set_breakpoints(self, file_path: str, 
                            breakpoints: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Set breakpoints in VBA code.
        
        Args:
            file_path: Source file path (module name or file)
            breakpoints: List of breakpoint specifications
            
        Returns:
            List of breakpoint results
        """
        # Extract module name from path
        # For VBA, the file path might be like "ModuleName.bas"
        module_name = file_path.replace('.bas', '').split('/')[-1]
        
        lines = [bp['line'] for bp in breakpoints]
        return await self.breakpoint_manager.set_breakpoints(module_name, lines)
    
    async def continue_execution(self):
        """Continue execution from current position."""
        logger.debug("Continuing execution")
        await self.com_bridge.execute('continue')
        self.is_paused = False
    
    async def step_over(self):
        """Step over current line."""
        logger.debug("Stepping over")
        await self.com_bridge.execute('step_over')
    
    async def step_in(self):
        """Step into function/sub."""
        logger.debug("Stepping in")
        await self.com_bridge.execute('step_in')
    
    async def step_out(self):
        """Step out of current function/sub."""
        logger.debug("Stepping out")
        await self.com_bridge.execute('step_out')
    
    async def pause(self):
        """Pause execution."""
        logger.debug("Pausing execution")
        # Send Ctrl+Break to VBA
        self.is_paused = True
    
    async def get_stack_trace(self) -> List[Dict[str, Any]]:
        """Get current call stack.
        
        Returns:
            List of stack frame dictionaries
        """
        try:
            frames = await self.com_bridge.execute('get_call_stack')
            
            # Convert to DAP format
            dap_frames = []
            for i, frame in enumerate(frames):
                dap_frames.append({
                    'id': i,
                    'name': frame.get('name', 'Unknown'),
                    'source': {
                        'name': frame.get('module', 'Unknown'),
                        'path': frame.get('module', 'Unknown') + '.bas',
                    },
                    'line': frame.get('line', 0),
                    'column': 0,
                })
            
            return dap_frames
            
        except Exception as e:
            logger.error(f"Failed to get stack trace: {e}")
            return []
    
    async def get_scopes(self, frame_id: int) -> List[Dict[str, Any]]:
        """Get variable scopes for a stack frame.
        
        Args:
            frame_id: Stack frame ID
            
        Returns:
            List of scope dictionaries
        """
        return [
            {
                'name': 'Locals',
                'variablesReference': frame_id * 1000 + 1,
                'expensive': False,
            },
            {
                'name': 'Globals',
                'variablesReference': frame_id * 1000 + 2,
                'expensive': False,
            },
        ]
    
    async def get_variables(self, variables_ref: int) -> List[Dict[str, Any]]:
        """Get variables for a scope.
        
        Args:
            variables_ref: Variables reference ID
            
        Returns:
            List of variable dictionaries
        """
        # This is a placeholder - actual implementation would query
        # VBA runtime through COM or watch window manipulation
        logger.warning("Variable inspection not fully implemented yet")
        return []
    
    async def evaluate(self, expression: str) -> Dict[str, Any]:
        """Evaluate an expression.
        
        Args:
            expression: Expression to evaluate
            
        Returns:
            Evaluation result dictionary
        """
        try:
            result = await self.com_bridge.execute('evaluate_expression', expression)
            return {
                'result': str(result),
                'variablesReference': 0,
            }
        except Exception as e:
            logger.error(f"Failed to evaluate expression: {e}")
            return {
                'result': f"Error: {e}",
                'variablesReference': 0,
            }
