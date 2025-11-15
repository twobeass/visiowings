"""Call stack inspection for VBA debugging.

Provides functionality to extract and format VBA call stack information.
"""

import logging
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)


class CallStackInspector:
    """Inspect VBA call stack during debugging.
    
    Extracts call stack frames from VBA debugger state and formats them
    for Debug Adapter Protocol.
    """

    def __init__(self, com_bridge):
        """Initialize call stack inspector.
        
        Args:
            com_bridge: COM bridge instance
        """
        self.com_bridge = com_bridge
        self._frame_id_counter = 0
    
    def get_stack_frames(self) -> List[Dict[str, any]]:
        """Get current call stack frames.
        
        Returns:
            List of stack frame dictionaries in DAP format
        """
        frames = []
        
        try:
            # Get current execution location (top of stack)
            current_frame = self._get_current_frame()
            if current_frame:
                frames.append(current_frame)
            
            # Try to get additional frames from call stack
            # Note: VBA COM API has limited call stack introspection
            # This is a best-effort implementation
            additional_frames = self._get_additional_frames()
            frames.extend(additional_frames)
            
        except Exception as e:
            logger.error(f"Failed to get stack frames: {e}", exc_info=True)
        
        return frames
    
    def _get_current_frame(self) -> Optional[Dict[str, any]]:
        """Get the current (top) stack frame.
        
        Returns:
            Current frame dictionary or None
        """
        try:
            if not self.com_bridge.vbe:
                return None
            
            vbe = self.com_bridge.vbe
            
            # Check if we have an active code pane (in break mode)
            if not hasattr(vbe, 'ActiveCodePane') or not vbe.ActiveCodePane:
                return None
            
            active_pane = vbe.ActiveCodePane
            code_module = active_pane.CodeModule
            
            # Get current line
            line = active_pane.TopLine if hasattr(active_pane, 'TopLine') else 1
            
            # Get module information
            module_name = 'Unknown'
            if hasattr(code_module, 'Parent'):
                module_name = code_module.Parent.Name
            
            # Try to get procedure name
            procedure_name = 'Unknown'
            try:
                if hasattr(code_module, 'ProcOfLine'):
                    procedure_name = code_module.ProcOfLine(line, 0)
            except:
                pass
            
            # Create frame
            frame_id = self._next_frame_id()
            return self._create_frame(
                frame_id=frame_id,
                name=procedure_name,
                module=module_name,
                line=line,
            )
            
        except Exception as e:
            logger.debug(f"Could not get current frame: {e}")
            return None
    
    def _get_additional_frames(self) -> List[Dict[str, any]]:
        """Get additional call stack frames.
        
        Returns:
            List of additional frames
            
        Note:
            VBA's COM API doesn't provide direct call stack access.
            This would require alternative approaches like:
            - Parsing error stack from runtime errors
            - Instrumenting code with stack tracking
            - Using Windows debugging APIs
        """
        # Placeholder for future implementation
        # Could be enhanced with:
        # - Analysis of VBA's call chain through debugger
        # - Integration with Windows debugging APIs
        # - Custom instrumentation
        return []
    
    def _create_frame(self, frame_id: int, name: str, 
                     module: str, line: int) -> Dict[str, any]:
        """Create a DAP-formatted stack frame.
        
        Args:
            frame_id: Unique frame identifier
            name: Procedure/function name
            module: Module name
            line: Line number
            
        Returns:
            DAP stack frame dictionary
        """
        return {
            'id': frame_id,
            'name': name,
            'source': {
                'name': f"{module}.bas",
                'path': f"{module}.bas",
            },
            'line': line,
            'column': 1,
        }
    
    def _next_frame_id(self) -> int:
        """Get next unique frame ID.
        
        Returns:
            Frame ID
        """
        self._frame_id_counter += 1
        return self._frame_id_counter
    
    def reset(self):
        """Reset frame ID counter."""
        self._frame_id_counter = 0
