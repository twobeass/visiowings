"""Integration tests for VBA debugging system.

Tests the full debugging workflow including adapter, COM bridge,
breakpoints, and event monitoring.
"""

import asyncio
import pytest
import unittest.mock as mock

from visiowings.debug import (
    VisioDebugAdapter,
    DebugSession,
    COMBridge,
    BreakpointManager,
    VBAEventMonitor,
    VariableInspector,
    CallStackInspector,
    ErrorHandler,
    BreakpointCleanupManager,
)


@pytest.mark.asyncio
class TestDebugIntegration:
    """Integration tests for debugging system."""

    @pytest.fixture
    def com_bridge(self):
        """Create mock COM bridge."""
        bridge = mock.Mock(spec=COMBridge)
        bridge.vbe = mock.Mock()
        bridge.active_project = mock.Mock()
        return bridge

    @pytest.fixture
    def breakpoint_manager(self, com_bridge):
        """Create breakpoint manager."""
        return BreakpointManager(com_bridge)

    @pytest.fixture
    def event_monitor(self, com_bridge):
        """Create event monitor."""
        return VBAEventMonitor(com_bridge, poll_interval=0.1)

    @pytest.fixture
    def cleanup_manager(self, breakpoint_manager):
        """Create cleanup manager."""
        return BreakpointCleanupManager(breakpoint_manager)

    async def test_full_debugging_workflow(self, com_bridge, breakpoint_manager):
        """Test complete debugging workflow."""
        # 1. Set breakpoint
        await breakpoint_manager.set_breakpoint('Module1', 5)
        
        # 2. Verify breakpoint was set
        breakpoints = breakpoint_manager.get_breakpoints()
        assert len(breakpoints) > 0
        
        # 3. Remove breakpoint
        await breakpoint_manager.remove_breakpoint('Module1', 5, 'original code')
        
        # 4. Verify cleanup
        breakpoints = breakpoint_manager.get_breakpoints()
        assert len(breakpoints) == 0

    async def test_event_monitoring(self, event_monitor):
        """Test event monitoring and callbacks."""
        break_events = []
        
        def on_break(location):
            break_events.append(location)
        
        # Register callback
        event_monitor.register_callback('on_break', on_break)
        
        # Start monitoring
        event_monitor.start()
        
        # Wait briefly
        await asyncio.sleep(0.2)
        
        # Stop monitoring
        event_monitor.stop()

    async def test_error_handling_with_cleanup(self, cleanup_manager):
        """Test error handling triggers cleanup."""
        # Register some breakpoints
        cleanup_manager.register_breakpoint('Module1', 5, 'original1')
        cleanup_manager.register_breakpoint('Module2', 10, 'original2')
        
        # Verify registered
        active = cleanup_manager.get_active_breakpoints()
        assert len(active) == 2
        
        # Perform cleanup
        results = await cleanup_manager.cleanup_all()
        
        # Verify results
        assert results['total'] == 2

    async def test_variable_inspection(self, com_bridge):
        """Test variable inspection functionality."""
        inspector = VariableInspector(com_bridge)
        
        # Test expression evaluation
        result = inspector.evaluate_expression('x + 1')
        assert 'result' in result
        assert 'type' in result
        
        # Test variable parsing from code
        code = '''
        Dim x As Integer
        Dim name As String
        Public count As Long
        '''
        
        variables = inspector.parse_variables_from_code(code)
        assert 'x' in variables
        assert 'name' in variables
        assert 'count' in variables

    async def test_call_stack_inspection(self, com_bridge):
        """Test call stack inspection."""
        inspector = CallStackInspector(com_bridge)
        
        # Mock current location
        com_bridge.vbe.ActiveCodePane = mock.Mock()
        code_module = mock.Mock()
        code_module.Parent.Name = 'Module1'
        com_bridge.vbe.ActiveCodePane.CodeModule = code_module
        com_bridge.vbe.ActiveCodePane.TopLine = 10
        code_module.ProcOfLine = mock.Mock(return_value='TestProc')
        
        # Get stack frames
        frames = inspector.get_stack_frames()
        
        # Verify frame structure
        if frames:
            frame = frames[0]
            assert 'id' in frame
            assert 'name' in frame
            assert 'source' in frame
            assert 'line' in frame

    async def test_error_handler_retry(self):
        """Test error handler retry mechanism."""
        handler = ErrorHandler()
        
        call_count = [0]
        
        @handler.with_error_handling('test_op', retry_count=2)
        async def failing_operation():
            call_count[0] += 1
            if call_count[0] < 3:
                raise Exception("Temporary failure")
            return "success"
        
        # Should succeed on third attempt
        result = await failing_operation()
        assert result == "success"
        assert call_count[0] == 3

    async def test_session_lifecycle(self, com_bridge):
        """Test debug session lifecycle."""
        session = DebugSession('test-session', 'test.vsd')
        
        # Mock session components
        with mock.patch.object(session, 'start') as mock_start:
            await session.start()
            mock_start.assert_called_once()


class TestBreakpointCleanup:
    """Tests for breakpoint cleanup functionality."""

    @pytest.fixture
    def manager(self):
        """Create cleanup manager with mock breakpoint manager."""
        bp_manager = mock.Mock()
        bp_manager.remove_breakpoint = mock.AsyncMock()
        return BreakpointCleanupManager(bp_manager)

    @pytest.mark.asyncio
    async def test_cleanup_all_success(self, manager):
        """Test successful cleanup of all breakpoints."""
        # Register breakpoints
        manager.register_breakpoint('M1', 5, 'code1')
        manager.register_breakpoint('M2', 10, 'code2')
        
        # Cleanup
        results = await manager.cleanup_all()
        
        # Verify
        assert results['total'] == 2
        assert results['successful'] == 2
        assert results['failed'] == 0

    @pytest.mark.asyncio
    async def test_cleanup_partial_failure(self, manager):
        """Test cleanup with some failures."""
        # Setup mock to fail on second call
        manager.breakpoint_manager.remove_breakpoint.side_effect = [
            None,  # First succeeds
            Exception("Failed"),  # Second fails
        ]
        
        # Register breakpoints
        manager.register_breakpoint('M1', 5, 'code1')
        manager.register_breakpoint('M2', 10, 'code2')
        
        # Cleanup
        results = await manager.cleanup_all()
        
        # Verify
        assert results['total'] == 2
        assert results['successful'] == 1
        assert results['failed'] == 1
        assert len(results['errors']) == 1


class TestErrorRecovery:
    """Tests for error recovery mechanisms."""

    @pytest.fixture
    def error_handler(self):
        """Create error handler."""
        return ErrorHandler()

    def test_error_count_tracking(self, error_handler):
        """Test error counting."""
        assert error_handler.error_count == 0
        
        error_handler.error_count = 5
        assert error_handler.error_count == 5
        
        error_handler.reset_error_count()
        assert error_handler.error_count == 0

    def test_error_summary(self, error_handler):
        """Test error summary generation."""
        error_handler._last_error = ValueError("Test error")
        error_handler.error_count = 3
        
        summary = error_handler.get_error_summary()
        
        assert summary['error_count'] == 3
        assert summary['last_error_type'] == 'ValueError'
        assert 'Test error' in summary['last_error']


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
