"""Call stack introspection for VBA debugging.

Provides call stack extraction and frame navigation for debugging sessions.
"""

import logging
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)


class CallStackInspector:
    """Inspector for VBA call stack.
    
    Extracts and provides access to the current call stack when
    VBA is in break mode.
    """

    def __init__(self, com_bridge):
        """Initialize call stack inspector.
        
        Args:
            com_bridge: COMBridge instance for VBA access
        """
        self.com_bridge = com_bridge
        self._frame_cache = {}
        self._next_frame_id = 1
        
    def get_stack_frames(self) -> List[Dict[str, any]]:
        """Get current call stack frames.
        
        Returns:
            List of stack frame dictionaries in DAP format
        """
        try:
            frames = []
            
            # Get current execution location
            current_location = self._get_current_location()
            
            if current_location:
                # Create frame for current location
                frame_id = self._get_frame_id(current_location)
                
                frame = {
                    'id': frame_id,
                    'name': current_location.get('procedure', '<main>'),
                    'source': {
                        'name': current_location.get('module', 'Unknown'),
                        'path': self._get_module_path(current_location.get('module')),
                    },
                    'line': current_location.get('line', 0),
                    'column': 0,
                }
                
                frames.append(frame)
            
            # Try to get parent frames (caller stack)
            # Note: VBA COM API has limited stack introspection
            # Full implementation would require more advanced techniques
            parent_frames = self._get_parent_frames()
            frames.extend(parent_frames)
            
            return frames
            
        except Exception as e:
            logger.error(f"Failed to get stack frames: {e}")
            return []
    
    def _get_current_location(self) -> Optional[Dict[str, any]]:
        """Get current execution location.
        
        Returns:
            Location dictionary or None
        """
        try:
            if not self.com_bridge.vbe:
                return None
            
            # Access active code pane
            active_pane = self.com_bridge.vbe.ActiveCodePane
            if not active_pane:
                return None
            
            location = {}
            
            # Get code module
            code_module = active_pane.CodeModule
            if code_module:
                # Get module name
                if hasattr(code_module, 'Parent'):
                    location['module'] = code_module.Parent.Name
                
                # Get current line
                try:
                    # TopLine gives the first visible line
                    # For break mode, this is typically the current execution line
                    location['line'] = active_pane.TopLine
                except:
                    pass
                
                # Get procedure name
                if location.get('line'):
                    try:
                        # ProcOfLine returns the procedure name at a given line
                        proc_name = code_module.ProcOfLine(location['line'], 0)
                        location['procedure'] = proc_name
                    except:
                        pass
            
            return location if location else None
            
        except Exception as e:
            logger.debug(f"Could not get current location: {e}")
            return None
    
    def _get_parent_frames(self) -> List[Dict[str, any]]:
        """Get parent/caller stack frames.
        
        Returns:
            List of parent frame dictionaries
        """
        # VBA's COM API doesn't provide direct stack access
        # This would require:
        # 1. Custom VBA instrumentation
        # 2. Parsing VBA call statements
        # 3. Using Windows debugging APIs
        
        # For now, return empty list
        # Future enhancement could implement stack walking
        return []
    
    def _get_frame_id(self, location: Dict[str, any]) -> int:
        """Get or create frame ID for a location.
        
        Args:
            location: Location dictionary
            
        Returns:
            Frame ID
        """
        key = f"{location.get('module')}:{location.get('line')}"
        
        if key not in self._frame_cache:
            self._frame_cache[key] = self._next_frame_id
            self._next_frame_id += 1
        
        return self._frame_cache[key]
    
    def _get_module_path(self, module_name: Optional[str]) -> str:
        """Get file path for a module.
        
        Args:
            module_name: Module name
            
        Returns:
            File path (may be synthetic)
        """
        if not module_name:
            return '<unknown>'
        
        # Try to map to exported file
        # In practice, this would need integration with visiowings export system
        return f"{module_name}.bas"
    
    def get_frame_scopes(self, frame_id: int) -> List[Dict[str, any]]:
        """Get variable scopes for a stack frame.
        
        Args:
            frame_id: Frame ID
            
        Returns:
            List of scope dictionaries
        """
        # Standard VBA scopes
        scopes = [
            {
                'name': 'Locals',
                'variablesReference': frame_id * 1000 + 1,
                'expensive': False,
            },
            {
                'name': 'Module',
                'variablesReference': frame_id * 1000 + 2,
                'expensive': False,
            },
            {
                'name': 'Global',
                'variablesReference': frame_id * 1000 + 3,
                'expensive': True,
            },
        ]
        
        return scopes
    
    def clear_cache(self):
        """Clear frame cache."""
        self._frame_cache.clear()
        self._next_frame_id = 1
        logger.debug("Call stack cache cleared")
    
    def get_frame_info(self, frame_id: int) -> Optional[Dict[str, any]]:
        """Get information about a specific frame.
        
        Args:
            frame_id: Frame ID
            
        Returns:
            Frame information or None
        """
        # Reverse lookup in cache
        for key, fid in self._frame_cache.items():
            if fid == frame_id:
                module, line = key.split(':', 1)
                return {
                    'module': module,
                    'line': int(line),
                    'frame_id': frame_id,
                }
        
        return None
