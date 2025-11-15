"""Tests for COM bridge."""

import asyncio
import pytest
from unittest.mock import MagicMock, patch

from visiowings.debug.com_bridge import COMBridge


@pytest.fixture
def bridge():
    """Create COM bridge instance."""
    return COMBridge()


def test_start_stop(bridge):
    """Test starting and stopping COM bridge."""
    bridge.start()
    assert bridge._running is True
    assert bridge._com_thread is not None
    
    bridge.stop()
    assert bridge._running is False


@pytest.mark.asyncio
@patch('win32com.client.Dispatch')
@patch('pythoncom.CoInitialize')
@patch('pythoncom.CoUninitialize')
async def test_connect_operation(mock_uninit, mock_init, mock_dispatch, bridge):
    """Test connect operation."""
    mock_app = MagicMock()
    mock_vbe = MagicMock()
    mock_app.VBE = mock_vbe
    mock_vbe.VBProjects.Count = 1
    mock_dispatch.return_value = mock_app
    
    bridge.start()
    
    # Direct test of operation (bypass async queue for testing)
    result = bridge._op_connect()
    
    assert result is True
    assert bridge.visio_app is not None
    mock_dispatch.assert_called_with('Visio.Application')
    
    bridge.stop()


@pytest.mark.asyncio
async def test_execute_timeout(bridge):
    """Test operation timeout."""
    bridge.start()
    
    with pytest.raises(TimeoutError):
        await bridge.execute('nonexistent_operation', timeout=0.1)
    
    bridge.stop()


@patch('win32com.client.Dispatch')
def test_get_modules_operation(mock_dispatch, bridge):
    """Test get modules operation."""
    mock_component = MagicMock()
    mock_component.Name = 'Module1'
    mock_component.Type = 1
    mock_component.CodeModule.CountOfLines = 100
    
    mock_project = MagicMock()
    mock_project.VBComponents = [mock_component]
    
    bridge.active_project = mock_project
    
    result = bridge._op_get_modules()
    
    assert len(result) == 1
    assert result[0]['name'] == 'Module1'
    assert result[0]['code_lines'] == 100


@patch('win32com.client.Dispatch')
def test_inject_breakpoint_operation(mock_dispatch, bridge):
    """Test breakpoint injection."""
    mock_code_module = MagicMock()
    mock_code_module.Lines.return_value = '    x = 10'
    
    mock_component = MagicMock()
    mock_component.CodeModule = mock_code_module
    
    mock_project = MagicMock()
    mock_project.VBComponents.return_value = mock_component
    
    bridge.active_project = mock_project
    
    original, line = bridge._op_inject_breakpoint('Module1', 5)
    
    assert original == '    x = 10'
    assert line == 5
    mock_code_module.ReplaceLine.assert_called_once()


@patch('win32com.client.Dispatch')
def test_remove_breakpoint_operation(mock_dispatch, bridge):
    """Test breakpoint removal."""
    mock_code_module = MagicMock()
    
    mock_component = MagicMock()
    mock_component.CodeModule = mock_code_module
    
    mock_project = MagicMock()
    mock_project.VBComponents.return_value = mock_component
    
    bridge.active_project = mock_project
    
    result = bridge._op_remove_breakpoint('Module1', 5, '    x = 10')
    
    assert result is True
    mock_code_module.ReplaceLine.assert_called_with(5, '    x = 10')


def test_get_debug_state(bridge):
    """Test get debug state operation."""
    mock_vbe = MagicMock()
    bridge.vbe = mock_vbe
    
    result = bridge._op_get_debug_state()
    
    assert 'mode' in result
    assert result['mode'] in ['design', 'run', 'break', 'unknown']
