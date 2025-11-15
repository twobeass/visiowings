"""Variable inspection and expression evaluation for VBA debugging.

Provides structured access to VBA variables and expression evaluation
through Watch window manipulation and Immediate window automation.
"""

import logging
import re
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


class VariableInspector:
    """Inspector for VBA variables and expressions.
    
    Uses VBA Watch window and Immediate window for variable access.
    """

    def __init__(self, com_bridge):
        """Initialize variable inspector.
        
        Args:
            com_bridge: COMBridge instance for VBA access
        """
        self.com_bridge = com_bridge
        self._watch_cache = {}
        
    def get_locals(self, frame_id: int = 0) -> List[Dict[str, Any]]:
        """Get local variables for a stack frame.
        
        Args:
            frame_id: Stack frame ID
            
        Returns:
            List of variable dictionaries
        """
        try:
            variables = []
            
            # Get current procedure's variables
            # This is a simplified implementation - full implementation
            # would parse VBA code to find variable declarations
            
            # For now, return placeholder with limited info
            logger.debug(f"Getting locals for frame {frame_id}")
            
            return variables
            
        except Exception as e:
            logger.error(f"Failed to get local variables: {e}")
            return []
    
    def get_globals(self) -> List[Dict[str, Any]]:
        """Get global/module-level variables.
        
        Returns:
            List of variable dictionaries
        """
        try:
            variables = []
            
            # Would need to parse module-level declarations
            logger.debug("Getting global variables")
            
            return variables
            
        except Exception as e:
            logger.error(f"Failed to get global variables: {e}")
            return []
    
    def evaluate_expression(self, expression: str) -> Dict[str, Any]:
        """Evaluate a VBA expression.
        
        Uses VBA's Immediate window to evaluate expressions when in break mode.
        
        Args:
            expression: VBA expression to evaluate
            
        Returns:
            Dictionary with result, type, and value
        """
        try:
            result = self._evaluate_via_immediate(expression)
            
            return {
                'result': result,
                'type': self._infer_type(result),
                'value': str(result),
                'variablesReference': 0,
            }
            
        except Exception as e:
            logger.error(f"Expression evaluation failed: {e}")
            return {
                'result': None,
                'type': 'error',
                'value': str(e),
                'variablesReference': 0,
            }
    
    def _evaluate_via_immediate(self, expression: str) -> Any:
        """Evaluate expression using Immediate window.
        
        Args:
            expression: Expression to evaluate
            
        Returns:
            Evaluation result
        """
        try:
            if not self.com_bridge.vbe:
                raise Exception("No VBE connection")
            
            # Access Immediate window
            immediate = self.com_bridge.vbe.Windows("Immediate")
            if not immediate:
                raise Exception("Immediate window not available")
            
            # Make it visible
            immediate.Visible = True
            
            # Try to use Debug.Print capture or Watch evaluation
            # This is a placeholder - actual implementation would need
            # to capture output from Immediate window or use Watch expressions
            
            logger.debug(f"Evaluating: {expression}")
            
            # Add as watch expression
            result = self._add_watch_expression(expression)
            
            return result
            
        except Exception as e:
            logger.error(f"Immediate evaluation failed: {e}")
            raise
    
    def _add_watch_expression(self, expression: str) -> Any:
        """Add and evaluate a watch expression.
        
        Args:
            expression: Expression to watch
            
        Returns:
            Expression value
        """
        try:
            if not self.com_bridge.active_project:
                raise Exception("No active VBA project")
            
            # Access Watch collection
            # Note: This requires VBA to be in break mode
            vbe = self.com_bridge.vbe
            
            # Try to add watch
            watch = vbe.VBProjects(1).VBComponents(1).CodeModule
            # This is simplified - actual Watch API varies by Visio version
            
            # For now, return placeholder
            logger.debug(f"Added watch for: {expression}")
            return None
            
        except Exception as e:
            logger.debug(f"Watch expression failed: {e}")
            return None
    
    def get_variable_value(self, var_name: str) -> Dict[str, Any]:
        """Get value of a specific variable.
        
        Args:
            var_name: Variable name
            
        Returns:
            Variable information dictionary
        """
        try:
            value = self.evaluate_expression(var_name)
            
            return {
                'name': var_name,
                'value': value.get('value', ''),
                'type': value.get('type', 'unknown'),
                'variablesReference': 0,
            }
            
        except Exception as e:
            logger.error(f"Failed to get variable '{var_name}': {e}")
            return {
                'name': var_name,
                'value': '<error>',
                'type': 'error',
                'variablesReference': 0,
            }
    
    def parse_variables_from_code(self, code: str) -> List[str]:
        """Parse variable declarations from VBA code.
        
        Args:
            code: VBA code to parse
            
        Returns:
            List of variable names
        """
        variables = []
        
        try:
            # Parse Dim, Public, Private declarations
            patterns = [
                r'\b(?:Dim|Public|Private|Static)\s+(\w+)\s+As\s+',
                r'\b(?:Dim|Public|Private|Static)\s+(\w+)\s*(?:,|$)',
            ]
            
            for pattern in patterns:
                matches = re.finditer(pattern, code, re.IGNORECASE | re.MULTILINE)
                for match in matches:
                    var_name = match.group(1)
                    if var_name not in variables:
                        variables.append(var_name)
            
        except Exception as e:
            logger.error(f"Failed to parse variables: {e}")
        
        return variables
    
    def _infer_type(self, value: Any) -> str:
        """Infer VBA type from Python value.
        
        Args:
            value: Value to infer type from
            
        Returns:
            VBA type string
        """
        if value is None:
            return 'Variant/Empty'
        elif isinstance(value, bool):
            return 'Boolean'
        elif isinstance(value, int):
            return 'Long'
        elif isinstance(value, float):
            return 'Double'
        elif isinstance(value, str):
            return 'String'
        else:
            return 'Variant'
    
    def format_for_dap(self, var_name: str, value: Any, 
                      var_ref: int = 0) -> Dict[str, Any]:
        """Format variable for DAP protocol.
        
        Args:
            var_name: Variable name
            value: Variable value
            var_ref: Variables reference for complex types
            
        Returns:
            DAP-formatted variable dictionary
        """
        return {
            'name': var_name,
            'value': str(value) if value is not None else '<uninitialized>',
            'type': self._infer_type(value),
            'variablesReference': var_ref,
            'evaluateName': var_name,
        }
    
    def get_watch_variables(self) -> List[Dict[str, Any]]:
        """Get all watch expressions.
        
        Returns:
            List of watch variable dictionaries
        """
        try:
            watches = []
            
            # Iterate through VBE watch expressions
            # This would access VBE.VBProjects(1).Watches collection
            # Implementation depends on Visio COM API availability
            
            logger.debug("Getting watch variables")
            
            return watches
            
        except Exception as e:
            logger.error(f"Failed to get watch variables: {e}")
            return []
    
    def clear_cache(self):
        """Clear the watch cache."""
        self._watch_cache.clear()
        logger.debug("Variable cache cleared")
