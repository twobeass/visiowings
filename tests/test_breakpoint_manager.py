"""Tests for breakpoint manager."""

import pytest
from unittest.mock import AsyncMock, MagicMock

from visiowings.debug.breakpoint_manager import Breakpoint, BreakpointManager


@pytest.fixture
def mock_bridge():
    """Create mock COM bridge."""
    bridge = MagicMock()
    bridge.execute = AsyncMock()
    return bridge


@pytest.fixture
def manager(mock_bridge):
    """Create breakpoint manager."""
    return BreakpointManager(mock_bridge)


def test_breakpoint_creation():
    """Test breakpoint object creation."""
    bp = Breakpoint('Module1', 10, verified=True)
    
    assert bp.module_name == 'Module1'
    assert bp.line_number == 10
    assert bp.verified is True
    assert bp.id == 'Module1:10'


@pytest.mark.asyncio
async def test_set_breakpoints_success(manager, mock_bridge):
    """Test successful breakpoint setting."""
    mock_bridge.execute.return_value = ('original line', 10)
    
    results = await manager.set_breakpoints('Module1', [10, 20])
    
    assert len(results) == 2
    assert all(r['verified'] for r in results)
    assert len(manager.breakpoints) == 2
    assert mock_bridge.execute.call_count == 2


@pytest.mark.asyncio
async def test_set_breakpoints_failure(manager, mock_bridge):
    """Test breakpoint setting with failures."""
    mock_bridge.execute.side_effect = [
        ('original', 10),
        Exception('Failed to inject')
    ]
    
    results = await manager.set_breakpoints('Module1', [10, 20])
    
    assert len(results) == 2
    assert results[0]['verified'] is True
    assert results[1]['verified'] is False
    assert 'message' in results[1]


@pytest.mark.asyncio
async def test_remove_breakpoint(manager, mock_bridge):
    """Test breakpoint removal."""
    # Setup breakpoint
    bp = Breakpoint('Module1', 10, verified=True)
    bp.original_line = 'x = 10'
    manager.breakpoints[bp.id] = bp
    
    mock_bridge.execute.return_value = True
    
    result = await manager.remove_breakpoint('Module1', 10)
    
    assert result is True
    assert len(manager.breakpoints) == 0
    mock_bridge.execute.assert_called_once()


@pytest.mark.asyncio
async def test_remove_nonexistent_breakpoint(manager):
    """Test removing breakpoint that doesn't exist."""
    result = await manager.remove_breakpoint('Module1', 10)
    
    assert result is False


@pytest.mark.asyncio
async def test_clear_module_breakpoints(manager, mock_bridge):
    """Test clearing all breakpoints in a module."""
    # Setup multiple breakpoints
    for line in [10, 20, 30]:
        bp = Breakpoint('Module1', line, verified=True)
        bp.original_line = f'line {line}'
        manager.breakpoints[bp.id] = bp
    
    # Add breakpoint in different module
    bp = Breakpoint('Module2', 5, verified=True)
    bp.original_line = 'line 5'
    manager.breakpoints[bp.id] = bp
    
    mock_bridge.execute.return_value = True
    
    await manager._clear_module_breakpoints('Module1')
    
    assert len(manager.breakpoints) == 1
    assert 'Module2:5' in manager.breakpoints


@pytest.mark.asyncio
async def test_clear_all_breakpoints(manager, mock_bridge):
    """Test clearing all breakpoints."""
    # Setup breakpoints
    for i, module in enumerate(['Module1', 'Module2', 'Module3']):
        bp = Breakpoint(module, 10, verified=True)
        bp.original_line = f'line {i}'
        manager.breakpoints[bp.id] = bp
    
    mock_bridge.execute.return_value = True
    
    await manager.clear_all()
    
    assert len(manager.breakpoints) == 0
    assert mock_bridge.execute.call_count == 3


def test_get_breakpoint(manager):
    """Test getting breakpoint by location."""
    bp = Breakpoint('Module1', 10, verified=True)
    manager.breakpoints[bp.id] = bp
    
    result = manager.get_breakpoint('Module1', 10)
    
    assert result is not None
    assert result.id == 'Module1:10'


def test_get_all_breakpoints(manager):
    """Test getting all breakpoints."""
    for line in [10, 20, 30]:
        bp = Breakpoint('Module1', line)
        manager.breakpoints[bp.id] = bp
    
    result = manager.get_all_breakpoints()
    
    assert len(result) == 3
    assert all(isinstance(bp, Breakpoint) for bp in result)
