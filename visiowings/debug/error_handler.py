"""Error handling and recovery for VBA debugging.

Provides robust error handling, retry logic, and recovery mechanisms.
"""

import asyncio
import functools
import logging
import traceback
from typing import Any, Callable, Optional

logger = logging.getLogger(__name__)


class ErrorHandler:
    """Handle errors and implement recovery strategies.
    
    Provides decorators and utilities for robust error handling
    with retry logic, timeouts, and error tracking.
    """

    def __init__(self):
        """Initialize error handler."""
        self.error_count = 0
        self._last_error: Optional[Exception] = None
        self._error_callbacks = []
    
    def reset_error_count(self):
        """Reset error counter."""
        self.error_count = 0
        logger.debug("Error count reset")
    
    def register_error_callback(self, callback: Callable[[Exception], None]):
        """Register a callback for error notifications.
        
        Args:
            callback: Function to call when errors occur
        """
        self._error_callbacks.append(callback)
    
    def _notify_error(self, error: Exception):
        """Notify registered callbacks of an error.
        
        Args:
            error: Exception that occurred
        """
        for callback in self._error_callbacks:
            try:
                callback(error)
            except Exception as e:
                logger.error(f"Error in error callback: {e}")
    
    def with_error_handling(self, operation_name: str, 
                           retry_count: int = 0,
                           timeout: Optional[float] = None):
        """Decorator for error handling with retry logic.
        
        Args:
            operation_name: Name of the operation for logging
            retry_count: Number of retry attempts
            timeout: Optional timeout in seconds
            
        Returns:
            Decorated function
        """
        def decorator(func: Callable) -> Callable:
            @functools.wraps(func)
            async def wrapper(*args, **kwargs) -> Any:
                last_error = None
                
                for attempt in range(retry_count + 1):
                    try:
                        # Apply timeout if specified
                        if timeout:
                            return await asyncio.wait_for(
                                func(*args, **kwargs),
                                timeout=timeout
                            )
                        else:
                            return await func(*args, **kwargs)
                            
                    except asyncio.TimeoutError:
                        last_error = TimeoutError(
                            f"{operation_name} timed out after {timeout}s"
                        )
                        logger.warning(f"Timeout in {operation_name}")
                        
                    except Exception as e:
                        last_error = e
                        logger.warning(
                            f"Error in {operation_name} (attempt {attempt + 1}/{retry_count + 1}): {e}"
                        )
                    
                    # Wait before retry
                    if attempt < retry_count:
                        await asyncio.sleep(0.5 * (attempt + 1))
                
                # All retries failed
                self.error_count += 1
                self._last_error = last_error
                self._notify_error(last_error)
                
                logger.error(
                    f"{operation_name} failed after {retry_count + 1} attempts: {last_error}"
                )
                raise last_error
                
            return wrapper
        return decorator
    
    def with_fallback(self, fallback_value: Any = None):
        """Decorator to return fallback value on error.
        
        Args:
            fallback_value: Value to return on error
            
        Returns:
            Decorated function
        """
        def decorator(func: Callable) -> Callable:
            @functools.wraps(func)
            async def wrapper(*args, **kwargs) -> Any:
                try:
                    return await func(*args, **kwargs)
                except Exception as e:
                    logger.warning(f"Function {func.__name__} failed, returning fallback: {e}")
                    self.error_count += 1
                    self._last_error = e
                    return fallback_value
            return wrapper
        return decorator
    
    def get_error_summary(self) -> dict:
        """Get summary of error state.
        
        Returns:
            Dictionary with error statistics
        """
        return {
            'error_count': self.error_count,
            'last_error': str(self._last_error) if self._last_error else None,
            'last_error_type': type(self._last_error).__name__ if self._last_error else None,
        }
    
    @staticmethod
    def log_exception(operation: str, error: Exception):
        """Log exception with full traceback.
        
        Args:
            operation: Operation that failed
            error: Exception that occurred
        """
        logger.error(
            f"Exception in {operation}: {error}\n"
            f"Traceback:\n{traceback.format_exc()}"
        )


class BreakpointCleanupManager:
    """Manage breakpoint cleanup on errors or shutdown.
    
    Ensures all injected breakpoints are removed even if
    debugging session terminates unexpectedly.
    """

    def __init__(self, breakpoint_manager):
        """Initialize cleanup manager.
        
        Args:
            breakpoint_manager: BreakpointManager instance
        """
        self.breakpoint_manager = breakpoint_manager
        self._active_breakpoints: List[tuple] = []
    
    def register_breakpoint(self, module: str, line: int, original_code: str):
        """Register a breakpoint for cleanup tracking.
        
        Args:
            module: Module name
            line: Line number
            original_code: Original code line
        """
        self._active_breakpoints.append((module, line, original_code))
        logger.debug(f"Registered breakpoint for cleanup: {module}:{line}")
    
    def unregister_breakpoint(self, module: str, line: int):
        """Unregister a breakpoint.
        
        Args:
            module: Module name
            line: Line number
        """
        self._active_breakpoints = [
            (m, l, c) for m, l, c in self._active_breakpoints
            if not (m == module and l == line)
        ]
        logger.debug(f"Unregistered breakpoint: {module}:{line}")
    
    def get_active_breakpoints(self) -> List[tuple]:
        """Get list of active breakpoints.
        
        Returns:
            List of (module, line, original_code) tuples
        """
        return self._active_breakpoints.copy()
    
    async def cleanup_all(self) -> dict:
        """Clean up all registered breakpoints.
        
        Returns:
            Cleanup results dictionary
        """
        results = {
            'total': len(self._active_breakpoints),
            'successful': 0,
            'failed': 0,
            'errors': [],
        }
        
        logger.info(f"Cleaning up {results['total']} breakpoints")
        
        for module, line, original_code in self._active_breakpoints:
            try:
                await self.breakpoint_manager.remove_breakpoint(
                    module, line, original_code
                )
                results['successful'] += 1
                logger.debug(f"Cleaned up breakpoint: {module}:{line}")
            except Exception as e:
                results['failed'] += 1
                error_msg = f"{module}:{line} - {str(e)}"
                results['errors'].append(error_msg)
                logger.error(f"Failed to cleanup breakpoint: {error_msg}")
        
        # Clear the list after cleanup attempt
        self._active_breakpoints.clear()
        
        logger.info(
            f"Cleanup complete: {results['successful']} successful, "
            f"{results['failed']} failed"
        )
        
        return results
