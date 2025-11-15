"""Error handling and recovery for VBA debugging.

Provides robust error handling, automatic cleanup, and recovery mechanisms
for debugging sessions.
"""

import asyncio
import functools
import logging
from typing import Any, Callable, Optional

logger = logging.getLogger(__name__)


class DebugError(Exception):
    """Base exception for debugging errors."""
    pass


class COMError(DebugError):
    """Error in COM communication."""
    pass


class BreakpointError(DebugError):
    """Error in breakpoint management."""
    pass


class SessionError(DebugError):
    """Error in debug session."""
    pass


class ErrorHandler:
    """Handles errors and provides recovery mechanisms."""

    def __init__(self, cleanup_callback: Optional[Callable] = None):
        """Initialize error handler.
        
        Args:
            cleanup_callback: Optional callback for cleanup operations
        """
        self.cleanup_callback = cleanup_callback
        self.error_count = 0
        self.max_errors = 10
        self._last_error = None
        
    def with_error_handling(self, operation_name: str, 
                           timeout: float = 5.0,
                           retry_count: int = 0,
                           cleanup_on_error: bool = False):
        """Decorator for error handling with timeout and retry.
        
        Args:
            operation_name: Name of operation for logging
            timeout: Operation timeout in seconds
            retry_count: Number of retries on failure
            cleanup_on_error: Whether to run cleanup on error
        """
        def decorator(func):
            @functools.wraps(func)
            async def async_wrapper(*args, **kwargs):
                last_error = None
                
                for attempt in range(retry_count + 1):
                    try:
                        # Execute with timeout
                        result = await asyncio.wait_for(
                            func(*args, **kwargs),
                            timeout=timeout
                        )
                        
                        # Reset error count on success
                        self.error_count = 0
                        return result
                        
                    except asyncio.TimeoutError as e:
                        last_error = e
                        logger.warning(
                            f"{operation_name} timed out after {timeout}s "
                            f"(attempt {attempt + 1}/{retry_count + 1})"
                        )
                        
                        if attempt < retry_count:
                            await asyncio.sleep(0.5 * (attempt + 1))
                        
                    except Exception as e:
                        last_error = e
                        logger.error(
                            f"{operation_name} failed: {e} "
                            f"(attempt {attempt + 1}/{retry_count + 1})",
                            exc_info=True
                        )
                        
                        if attempt < retry_count:
                            await asyncio.sleep(0.5 * (attempt + 1))
                
                # All attempts failed
                self.error_count += 1
                self._last_error = last_error
                
                if cleanup_on_error and self.cleanup_callback:
                    try:
                        await self.cleanup_callback()
                    except Exception as cleanup_err:
                        logger.error(f"Cleanup failed: {cleanup_err}")
                
                # Check if we've hit max errors
                if self.error_count >= self.max_errors:
                    logger.critical(
                        f"Max errors ({self.max_errors}) reached. "
                        "Halting operations."
                    )
                    raise SessionError("Too many errors - session terminated")
                
                raise DebugError(
                    f"{operation_name} failed after {retry_count + 1} attempts: {last_error}"
                )
            
            @functools.wraps(func)
            def sync_wrapper(*args, **kwargs):
                try:
                    result = func(*args, **kwargs)
                    self.error_count = 0
                    return result
                except Exception as e:
                    self.error_count += 1
                    self._last_error = e
                    
                    logger.error(
                        f"{operation_name} failed: {e}",
                        exc_info=True
                    )
                    
                    if cleanup_on_error and self.cleanup_callback:
                        try:
                            self.cleanup_callback()
                        except Exception as cleanup_err:
                            logger.error(f"Cleanup failed: {cleanup_err}")
                    
                    if self.error_count >= self.max_errors:
                        raise SessionError("Too many errors - session terminated")
                    
                    raise DebugError(f"{operation_name} failed: {e}")
            
            # Return appropriate wrapper based on function type
            if asyncio.iscoroutinefunction(func):
                return async_wrapper
            else:
                return sync_wrapper
        
        return decorator
    
    def reset_error_count(self):
        """Reset the error counter."""
        self.error_count = 0
        self._last_error = None
        logger.debug("Error count reset")
    
    def get_last_error(self) -> Optional[Exception]:
        """Get the last error that occurred.
        
        Returns:
            Last exception or None
        """
        return self._last_error
    
    def get_error_summary(self) -> dict:
        """Get error summary.
        
        Returns:
            Dictionary with error information
        """
        return {
            'error_count': self.error_count,
            'max_errors': self.max_errors,
            'last_error': str(self._last_error) if self._last_error else None,
            'last_error_type': type(self._last_error).__name__ if self._last_error else None,
        }


class BreakpointCleanupManager:
    """Manages automatic cleanup of injected breakpoints."""

    def __init__(self, breakpoint_manager):
        """Initialize cleanup manager.
        
        Args:
            breakpoint_manager: BreakpointManager instance
        """
        self.breakpoint_manager = breakpoint_manager
        self._active_breakpoints = {}
        
    def register_breakpoint(self, module: str, line: int, original_code: str):
        """Register a breakpoint for cleanup.
        
        Args:
            module: Module name
            line: Line number
            original_code: Original code to restore
        """
        key = f"{module}:{line}"
        self._active_breakpoints[key] = {
            'module': module,
            'line': line,
            'original': original_code,
        }
        logger.debug(f"Registered breakpoint for cleanup: {key}")
    
    def unregister_breakpoint(self, module: str, line: int):
        """Unregister a breakpoint.
        
        Args:
            module: Module name
            line: Line number
        """
        key = f"{module}:{line}"
        if key in self._active_breakpoints:
            del self._active_breakpoints[key]
            logger.debug(f"Unregistered breakpoint: {key}")
    
    async def cleanup_all(self) -> dict:
        """Clean up all registered breakpoints.
        
        Returns:
            Dictionary with cleanup results
        """
        results = {
            'total': len(self._active_breakpoints),
            'successful': 0,
            'failed': 0,
            'errors': [],
        }
        
        if not self._active_breakpoints:
            logger.info("No breakpoints to clean up")
            return results
        
        logger.info(f"Cleaning up {len(self._active_breakpoints)} breakpoints")
        
        for key, bp_info in list(self._active_breakpoints.items()):
            try:
                await self.breakpoint_manager.remove_breakpoint(
                    bp_info['module'],
                    bp_info['line'],
                    bp_info['original']
                )
                results['successful'] += 1
                self.unregister_breakpoint(bp_info['module'], bp_info['line'])
                
            except Exception as e:
                results['failed'] += 1
                results['errors'].append({
                    'breakpoint': key,
                    'error': str(e),
                })
                logger.error(f"Failed to cleanup breakpoint {key}: {e}")
        
        logger.info(
            f"Cleanup complete: {results['successful']} successful, "
            f"{results['failed']} failed"
        )
        
        return results
    
    def get_active_breakpoints(self) -> list:
        """Get list of active breakpoints.
        
        Returns:
            List of breakpoint dictionaries
        """
        return list(self._active_breakpoints.values())


class RecoveryManager:
    """Manages recovery from errors."""

    def __init__(self, com_bridge, cleanup_manager: BreakpointCleanupManager):
        """Initialize recovery manager.
        
        Args:
            com_bridge: COMBridge instance
            cleanup_manager: BreakpointCleanupManager instance
        """
        self.com_bridge = com_bridge
        self.cleanup_manager = cleanup_manager
        
    async def recover_from_crash(self) -> bool:
        """Attempt to recover from a crash.
        
        Returns:
            True if recovery successful
        """
        logger.info("Attempting crash recovery...")
        
        try:
            # 1. Clean up all breakpoints
            cleanup_results = await self.cleanup_manager.cleanup_all()
            
            if cleanup_results['failed'] > 0:
                logger.warning(
                    f"Some breakpoints could not be cleaned: "
                    f"{cleanup_results['failed']}"
                )
            
            # 2. Try to reconnect to Visio
            # (COM bridge would need reconnect method)
            
            logger.info("Recovery completed")
            return True
            
        except Exception as e:
            logger.error(f"Recovery failed: {e}", exc_info=True)
            return False
    
    async def safe_shutdown(self):
        """Perform safe shutdown with cleanup."""
        logger.info("Performing safe shutdown...")
        
        try:
            # Clean up breakpoints
            await self.cleanup_manager.cleanup_all()
            
            # Stop COM bridge
            if self.com_bridge:
                self.com_bridge.stop()
            
            logger.info("Safe shutdown complete")
            
        except Exception as e:
            logger.error(f"Error during shutdown: {e}", exc_info=True)
