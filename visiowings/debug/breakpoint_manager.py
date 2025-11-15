"""Breakpoint management for VBA debugging.

Handles injection, removal, and tracking of breakpoints in VBA code.
"""

import logging
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


class Breakpoint:
    """Represents a breakpoint in VBA code."""
    
    def __init__(self, module_name: str, line_number: int, verified: bool = False):
        self.module_name = module_name
        self.line_number = line_number
        self.verified = verified
        self.original_line: Optional[str] = None
        self.id = f"{module_name}:{line_number}"
    
    def __repr__(self):
        return f"<Breakpoint {self.id} verified={self.verified}>"


class BreakpointManager:
    """Manages breakpoints for VBA debugging.
    
    Handles:
    - Breakpoint injection into VBA code
    - Original code preservation
    - Breakpoint removal and restoration
    - Edge case handling (locked modules, existing stops, etc.)
    """
    
    def __init__(self, com_bridge):
        """Initialize breakpoint manager.
        
        Args:
            com_bridge: COMBridge instance for VBA operations
        """
        self.com_bridge = com_bridge
        self.breakpoints: Dict[str, Breakpoint] = {}
        
    async def set_breakpoints(self, module_name: str, 
                            lines: List[int]) -> List[Dict[str, any]]:
        """Set breakpoints at specified lines.
        
        Args:
            module_name: VBA module name
            lines: List of line numbers for breakpoints
            
        Returns:
            List of breakpoint result dictionaries
        """
        results = []
        
        # Remove existing breakpoints for this module
        await self._clear_module_breakpoints(module_name)
        
        # Set new breakpoints
        for line in lines:
            try:
                original_line, actual_line = await self.com_bridge.execute(
                    'inject_breakpoint', module_name, line
                )
                
                bp = Breakpoint(module_name, actual_line, verified=True)
                bp.original_line = original_line
                self.breakpoints[bp.id] = bp
                
                results.append({
                    'verified': True,
                    'line': actual_line,
                    'id': len(results),
                })
                
                logger.info(f"Set breakpoint at {bp.id}")
                
            except Exception as e:
                logger.error(f"Failed to set breakpoint at {module_name}:{line}: {e}")
                results.append({
                    'verified': False,
                    'line': line,
                    'message': str(e),
                })
        
        return results
    
    async def remove_breakpoint(self, module_name: str, line_number: int) -> bool:
        """Remove a specific breakpoint.
        
        Args:
            module_name: Module name
            line_number: Line number
            
        Returns:
            True if successful
        """
        bp_id = f"{module_name}:{line_number}"
        bp = self.breakpoints.get(bp_id)
        
        if not bp:
            logger.warning(f"Breakpoint {bp_id} not found")
            return False
        
        try:
            if bp.original_line:
                success = await self.com_bridge.execute(
                    'remove_breakpoint', 
                    module_name, 
                    line_number, 
                    bp.original_line
                )
                
                if success:
                    del self.breakpoints[bp_id]
                    logger.info(f"Removed breakpoint {bp_id}")
                    return True
            
            return False
            
        except Exception as e:
            logger.error(f"Failed to remove breakpoint {bp_id}: {e}")
            return False
    
    async def _clear_module_breakpoints(self, module_name: str):
        """Clear all breakpoints in a module.
        
        Args:
            module_name: Module name
        """
        to_remove = [
            bp_id for bp_id, bp in self.breakpoints.items()
            if bp.module_name == module_name
        ]
        
        for bp_id in to_remove:
            bp = self.breakpoints[bp_id]
            await self.remove_breakpoint(bp.module_name, bp.line_number)
    
    async def clear_all(self):
        """Clear all breakpoints."""
        bp_ids = list(self.breakpoints.keys())
        for bp_id in bp_ids:
            bp = self.breakpoints[bp_id]
            await self.remove_breakpoint(bp.module_name, bp.line_number)
        
        logger.info("Cleared all breakpoints")
    
    def get_breakpoint(self, module_name: str, line_number: int) -> Optional[Breakpoint]:
        """Get breakpoint by location.
        
        Args:
            module_name: Module name
            line_number: Line number
            
        Returns:
            Breakpoint if found, None otherwise
        """
        bp_id = f"{module_name}:{line_number}"
        return self.breakpoints.get(bp_id)
    
    def get_all_breakpoints(self) -> List[Breakpoint]:
        """Get all breakpoints.
        
        Returns:
            List of all breakpoints
        """
        return list(self.breakpoints.values())
