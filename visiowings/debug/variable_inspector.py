"""Variable inspection for VBA debugging.

Provides functionality to inspect and evaluate VBA variables
during debugging sessions.
"""

import logging
import re
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


class VariableInspector:
    """Inspect VBA variables during debugging.
    
    Provides methods to evaluate expressions, extract variable values,
    and parse variable declarations from code.
    """

    def __init__(self, com_bridge):
        """Initialize variable inspector.
        
        Args:
            com_bridge: COM bridge instance
        """
        self.com_bridge = com_bridge
        self._watch_cache: Dict[str, Any] = {}
    
    def evaluate_expression(self, expression: str) -> Dict[str, Any]:
        """Evaluate a VBA expression.
        
        Args:
            expression: VBA expression to evaluate
            
        Returns:
            Dictionary with result, type, and success status
        """
        try:
            # Try to evaluate using VBA's immediate window or watch
            result = self._evaluate_via_watch(expression)
            
            return {
                'success': True,
                'result': str(result) if result is not None else 'Nothing',
                'type': self._infer_type(result),
            }
        except Exception as e:
            logger.error(f"Failed to evaluate '{expression}': {e}")
            return {
                'success': False,
                'result': None,
                'error': str(e),
                'type': 'unknown',
            }
    
    def _evaluate_via_watch(self, expression: str) -> Any:
        """Evaluate expression via VBA watch window.
        
        Args:
            expression: Expression to evaluate
            
        Returns:
            Evaluation result
        """
        if not self.com_bridge.vbe:
            raise Exception("No VBE connection")
        
        try:
            # Access main window and watch window
            vbe = self.com_bridge.vbe
            
            # Try to use Immediate window for evaluation
            # Note: This is a simplified implementation
            # Full implementation would need to interact with VBA's evaluation context
            
            # For now, cache and return placeholder
            self._watch_cache[expression] = None
            return None
            
        except Exception as e:
            logger.debug(f"Watch evaluation failed: {e}")
            raise
    
    def _infer_type(self, value: Any) -> str:
        """Infer VBA type from Python value.
        
        Args:
            value: Value to inspect
            
        Returns:
            VBA type name
        """
        if value is None:
            return 'Variant'
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
    
    def parse_variables_from_code(self, code: str) -> Dict[str, Dict[str, str]]:
        """Parse variable declarations from VBA code.
        
        Args:
            code: VBA code to parse
            
        Returns:
            Dictionary mapping variable names to their attributes
        """
        variables = {}
        
        # Patterns for variable declarations
        patterns = [
            # Dim x As Type
            r'Dim\s+(\w+)\s+As\s+(\w+)',
            # Public/Private x As Type  
            r'(Public|Private)\s+(\w+)\s+As\s+(\w+)',
            # Static x As Type
            r'Static\s+(\w+)\s+As\s+(\w+)',
        ]
        
        for pattern in patterns:
            matches = re.finditer(pattern, code, re.IGNORECASE)
            for match in matches:
                groups = match.groups()
                if len(groups) == 2:
                    # Dim var As Type
                    name, vba_type = groups
                    scope = 'local'
                elif len(groups) == 3:
                    # Public/Private var As Type
                    scope, name, vba_type = groups
                    scope = scope.lower()
                else:
                    continue
                
                variables[name] = {
                    'type': vba_type,
                    'scope': scope,
                }
        
        return variables
    
    def get_local_variables(self, module_name: str, 
                           procedure_name: str) -> List[Dict[str, Any]]:
        """Get local variables for a procedure.
        
        Args:
            module_name: Module name
            procedure_name: Procedure name
            
        Returns:
            List of variable dictionaries
        """
        try:
            # Get procedure code
            code = self._get_procedure_code(module_name, procedure_name)
            if not code:
                return []
            
            # Parse variables
            variables = self.parse_variables_from_code(code)
            
            # Convert to DAP format
            result = []
            for name, attrs in variables.items():
                result.append({
                    'name': name,
                    'value': '<unknown>',
                    'type': attrs.get('type', 'Variant'),
                    'variablesReference': 0,
                })
            
            return result
            
        except Exception as e:
            logger.error(f"Failed to get local variables: {e}")
            return []
    
    def _get_procedure_code(self, module_name: str, 
                           procedure_name: str) -> Optional[str]:
        """Get code for a specific procedure.
        
        Args:
            module_name: Module name
            procedure_name: Procedure name
            
        Returns:
            Procedure code or None
        """
        try:
            if not self.com_bridge.active_project:
                return None
            
            component = self.com_bridge.active_project.VBComponents(module_name)
            code_module = component.CodeModule
            
            # Find procedure start and end lines
            proc_start_line = code_module.ProcStartLine(procedure_name, 0)
            proc_line_count = code_module.ProcCountLines(procedure_name, 0)
            
            # Get procedure code
            return code_module.Lines(proc_start_line, proc_line_count)
            
        except Exception as e:
            logger.debug(f"Could not get procedure code: {e}")
            return None
    
    def format_for_dap(self, variables: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Format variables for Debug Adapter Protocol.
        
        Args:
            variables: Dictionary of variables
            
        Returns:
            List of DAP-formatted variable dictionaries
        """
        result = []
        
        for name, value in variables.items():
            result.append({
                'name': name,
                'value': str(value) if value is not None else 'Nothing',
                'type': self._infer_type(value),
                'variablesReference': 0,
            })
        
        return result
