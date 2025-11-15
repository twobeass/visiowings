"""Tests for debug adapter."""

import asyncio
import json
import pytest
from unittest.mock import AsyncMock, MagicMock, patch

from visiowings.debug.debug_adapter import VisioDebugAdapter


class MockStreamReader:
    def __init__(self, messages):
        self.messages = messages
        self.index = 0
    
    async def readline(self):
        if self.index >= len(self.messages):
            return b''
        msg = self.messages[self.index]
        self.index += 1
        return msg
    
    async def readexactly(self, n):
        if self.index >= len(self.messages):
            return b''
        msg = self.messages[self.index]
        self.index += 1
        return msg


class MockStreamWriter:
    def __init__(self):
        self.data = []
    
    def write(self, data):
        self.data.append(data)
    
    async def drain(self):
        pass
    
    def close(self):
        pass
    
    async def wait_closed(self):
        pass


@pytest.fixture
def adapter():
    return VisioDebugAdapter()


@pytest.mark.asyncio
async def test_initialize_request(adapter):
    """Test initialize request handling."""
    message = {
        'type': 'request',
        'seq': 1,
        'command': 'initialize',
        'arguments': {}
    }
    
    response = await adapter.handle_request('initialize', message)
    
    assert response['success'] is True
    assert 'supportsConfigurationDoneRequest' in response['body']
    assert response['body']['supportsConfigurationDoneRequest'] is True


@pytest.mark.asyncio
async def test_launch_request_missing_file(adapter):
    """Test launch request with missing visio file."""
    message = {
        'type': 'request',
        'seq': 2,
        'command': 'launch',
        'arguments': {}
    }
    
    response = await adapter.handle_request('launch', message)
    
    assert response['success'] is False
    assert 'required' in response['message'].lower()


@pytest.mark.asyncio
@patch('visiowings.debug.debug_session.DebugSession')
async def test_launch_request_success(mock_session_class, adapter):
    """Test successful launch request."""
    mock_session = AsyncMock()
    mock_session.start = AsyncMock()
    mock_session_class.return_value = mock_session
    
    message = {
        'type': 'request',
        'seq': 2,
        'command': 'launch',
        'arguments': {
            'visioFile': 'test.vsd'
        }
    }
    
    response = await adapter.handle_request('launch', message)
    
    assert response['success'] is True
    assert len(adapter.sessions) == 1


@pytest.mark.asyncio
async def test_set_breakpoints(adapter):
    """Test set breakpoints request."""
    # Setup mock session
    mock_session = AsyncMock()
    mock_session.set_breakpoints = AsyncMock(return_value=[
        {'verified': True, 'line': 5, 'id': 0}
    ])
    adapter.sessions['test'] = mock_session
    
    message = {
        'type': 'request',
        'seq': 3,
        'command': 'setBreakpoints',
        'arguments': {
            'source': {'path': 'Module1.bas'},
            'breakpoints': [{'line': 5}]
        }
    }
    
    response = await adapter.handle_request('setBreakpoints', message)
    
    assert response['success'] is True
    assert len(response['body']['breakpoints']) == 1
    assert response['body']['breakpoints'][0]['verified'] is True


@pytest.mark.asyncio
async def test_continue_request(adapter):
    """Test continue request."""
    mock_session = AsyncMock()
    mock_session.continue_execution = AsyncMock()
    adapter.sessions['test'] = mock_session
    
    message = {
        'type': 'request',
        'seq': 4,
        'command': 'continue'
    }
    
    response = await adapter.handle_request('continue', message)
    
    assert response['success'] is True
    mock_session.continue_execution.assert_called_once()


@pytest.mark.asyncio
async def test_step_commands(adapter):
    """Test step over, in, out commands."""
    mock_session = AsyncMock()
    mock_session.step_over = AsyncMock()
    mock_session.step_in = AsyncMock()
    mock_session.step_out = AsyncMock()
    adapter.sessions['test'] = mock_session
    
    # Step over
    response = await adapter.handle_request('next', {'seq': 5, 'command': 'next'})
    assert response['success'] is True
    mock_session.step_over.assert_called_once()
    
    # Step in
    response = await adapter.handle_request('stepIn', {'seq': 6, 'command': 'stepIn'})
    assert response['success'] is True
    mock_session.step_in.assert_called_once()
    
    # Step out
    response = await adapter.handle_request('stepOut', {'seq': 7, 'command': 'stepOut'})
    assert response['success'] is True
    mock_session.step_out.assert_called_once()


@pytest.mark.asyncio
async def test_stack_trace(adapter):
    """Test stack trace request."""
    mock_session = AsyncMock()
    mock_session.get_stack_trace = AsyncMock(return_value=[
        {'id': 0, 'name': 'TestFunc', 'line': 10}
    ])
    adapter.sessions['test'] = mock_session
    
    message = {
        'type': 'request',
        'seq': 8,
        'command': 'stackTrace'
    }
    
    response = await adapter.handle_request('stackTrace', message)
    
    assert response['success'] is True
    assert len(response['body']['stackFrames']) == 1


@pytest.mark.asyncio
async def test_disconnect(adapter):
    """Test disconnect request."""
    mock_session = AsyncMock()
    mock_session.disconnect = AsyncMock()
    adapter.sessions['test'] = mock_session
    
    message = {
        'type': 'request',
        'seq': 9,
        'command': 'disconnect'
    }
    
    response = await adapter.handle_request('disconnect', message)
    
    assert response['success'] is True
    assert len(adapter.sessions) == 0
    mock_session.disconnect.assert_called_once()


def test_create_response(adapter):
    """Test response creation."""
    request = {'seq': 10, 'command': 'test'}
    body = {'result': 'success'}
    
    response = adapter.create_response(request, body)
    
    assert response['type'] == 'response'
    assert response['success'] is True
    assert response['command'] == 'test'
    assert response['body'] == body


def test_create_error_response(adapter):
    """Test error response creation."""
    request = {'seq': 11, 'command': 'test'}
    error_msg = 'Test error'
    
    response = adapter.create_error_response(request, error_msg)
    
    assert response['type'] == 'response'
    assert response['success'] is False
    assert response['message'] == error_msg
