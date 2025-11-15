"""Debug Adapter Protocol (DAP) implementation for Visio VBA.

Provides VS Code debug adapter that communicates with the COM bridge
to enable remote debugging of VBA code in Visio.
"""

import asyncio
import json
import logging
import sys
from typing import Any, Dict, Optional

from .debug_session import DebugSession

logger = logging.getLogger(__name__)


class VisioDebugAdapter:
    """Debug Adapter Protocol server for Visio VBA debugging.
    
    Implements DAP to enable VS Code integration with Visio VBA debugging.
    Handles multiple concurrent sessions with proper state management.
    """

    def __init__(self):
        self.sessions: Dict[str, DebugSession] = {}
        self.sequence = 1
        self.reader: Optional[asyncio.StreamReader] = None
        self.writer: Optional[asyncio.StreamWriter] = None
        
    async def start_server(self, host: str = '127.0.0.1', port: int = 5678):
        """Start the DAP server.
        
        Args:
            host: Host address to bind to
            port: Port number to listen on
        """
        server = await asyncio.start_server(
            self.handle_client, host, port
        )
        logger.info(f"Debug adapter server started on {host}:{port}")
        
        async with server:
            await server.serve_forever()
    
    async def handle_client(self, reader: asyncio.StreamReader, 
                           writer: asyncio.StreamWriter):
        """Handle incoming client connection.
        
        Args:
            reader: Stream reader for incoming messages
            writer: Stream writer for outgoing messages
        """
        self.reader = reader
        self.writer = writer
        
        try:
            while True:
                message = await self.read_message()
                if message is None:
                    break
                    
                response = await self.handle_message(message)
                if response:
                    await self.send_message(response)
                    
        except Exception as e:
            logger.error(f"Error handling client: {e}", exc_info=True)
        finally:
            writer.close()
            await writer.wait_closed()
    
    async def read_message(self) -> Optional[Dict[str, Any]]:
        """Read a DAP message from the input stream.
        
        Returns:
            Parsed JSON message or None if connection closed
        """
        if not self.reader:
            return None
            
        try:
            # Read headers
            headers = {}
            while True:
                line = await self.reader.readline()
                if not line or line == b'\r\n':
                    break
                key, value = line.decode('utf-8').strip().split(': ', 1)
                headers[key] = value
            
            # Read content
            content_length = int(headers.get('Content-Length', 0))
            if content_length == 0:
                return None
                
            content = await self.reader.readexactly(content_length)
            return json.loads(content.decode('utf-8'))
            
        except Exception as e:
            logger.error(f"Error reading message: {e}")
            return None
    
    async def send_message(self, message: Dict[str, Any]):
        """Send a DAP message to the output stream.
        
        Args:
            message: Message dictionary to send
        """
        if not self.writer:
            return
            
        try:
            content = json.dumps(message).encode('utf-8')
            headers = f"Content-Length: {len(content)}\r\n\r\n".encode('utf-8')
            
            self.writer.write(headers + content)
            await self.writer.drain()
            
        except Exception as e:
            logger.error(f"Error sending message: {e}")
    
    async def handle_message(self, message: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Handle incoming DAP message.
        
        Args:
            message: Incoming message dictionary
            
        Returns:
            Response message or None
        """
        msg_type = message.get('type')
        command = message.get('command')
        
        if msg_type == 'request':
            return await self.handle_request(command, message)
        
        return None
    
    async def handle_request(self, command: str, 
                           message: Dict[str, Any]) -> Dict[str, Any]:
        """Handle DAP request.
        
        Args:
            command: Command name
            message: Request message
            
        Returns:
            Response message
        """
        handlers = {
            'initialize': self.handle_initialize,
            'launch': self.handle_launch,
            'attach': self.handle_attach,
            'setBreakpoints': self.handle_set_breakpoints,
            'continue': self.handle_continue,
            'next': self.handle_next,
            'stepIn': self.handle_step_in,
            'stepOut': self.handle_step_out,
            'pause': self.handle_pause,
            'stackTrace': self.handle_stack_trace,
            'scopes': self.handle_scopes,
            'variables': self.handle_variables,
            'evaluate': self.handle_evaluate,
            'disconnect': self.handle_disconnect,
        }
        
        handler = handlers.get(command)
        if handler:
            try:
                return await handler(message)
            except Exception as e:
                logger.error(f"Error handling {command}: {e}", exc_info=True)
                return self.create_error_response(message, str(e))
        
        return self.create_error_response(message, f"Unknown command: {command}")
    
    async def handle_initialize(self, message: Dict[str, Any]) -> Dict[str, Any]:
        """Handle initialize request."""
        return self.create_response(message, {
            'supportsConfigurationDoneRequest': True,
            'supportsEvaluateForHovers': True,
            'supportsStepBack': False,
            'supportsSetVariable': False,
            'supportsRestartFrame': False,
            'supportsConditionalBreakpoints': False,
            'supportsHitConditionalBreakpoints': False,
        })
    
    async def handle_launch(self, message: Dict[str, Any]) -> Dict[str, Any]:
        """Handle launch request."""
        args = message.get('arguments', {})
        visio_file = args.get('visioFile')
        
        if not visio_file:
            return self.create_error_response(message, "visioFile is required")
        
        session_id = f"session_{len(self.sessions)}"
        session = DebugSession(session_id, visio_file)
        self.sessions[session_id] = session
        
        try:
            await session.start()
            await self.send_event('initialized', {})
            return self.create_response(message, {})
        except Exception as e:
            return self.create_error_response(message, str(e))
    
    async def handle_attach(self, message: Dict[str, Any]) -> Dict[str, Any]:
        """Handle attach request."""
        args = message.get('arguments', {})
        session_id = args.get('sessionId')
        
        if session_id and session_id in self.sessions:
            session = self.sessions[session_id]
            await session.reconnect()
        else:
            # Create new attach session
            visio_file = args.get('visioFile')
            session = DebugSession(session_id or f"session_{len(self.sessions)}", visio_file)
            self.sessions[session.session_id] = session
            await session.attach()
        
        await self.send_event('initialized', {})
        return self.create_response(message, {})
    
    async def handle_set_breakpoints(self, message: Dict[str, Any]) -> Dict[str, Any]:
        """Handle setBreakpoints request."""
        args = message.get('arguments', {})
        source = args.get('source', {})
        breakpoints = args.get('breakpoints', [])
        
        # Get active session (use first for now)
        if not self.sessions:
            return self.create_error_response(message, "No active debug session")
        
        session = list(self.sessions.values())[0]
        result_bps = await session.set_breakpoints(source.get('path'), breakpoints)
        
        return self.create_response(message, {'breakpoints': result_bps})
    
    async def handle_continue(self, message: Dict[str, Any]) -> Dict[str, Any]:
        """Handle continue request."""
        session = list(self.sessions.values())[0] if self.sessions else None
        if session:
            await session.continue_execution()
        return self.create_response(message, {'allThreadsContinued': True})
    
    async def handle_next(self, message: Dict[str, Any]) -> Dict[str, Any]:
        """Handle next (step over) request."""
        session = list(self.sessions.values())[0] if self.sessions else None
        if session:
            await session.step_over()
        return self.create_response(message, {})
    
    async def handle_step_in(self, message: Dict[str, Any]) -> Dict[str, Any]:
        """Handle stepIn request."""
        session = list(self.sessions.values())[0] if self.sessions else None
        if session:
            await session.step_in()
        return self.create_response(message, {})
    
    async def handle_step_out(self, message: Dict[str, Any]) -> Dict[str, Any]:
        """Handle stepOut request."""
        session = list(self.sessions.values())[0] if self.sessions else None
        if session:
            await session.step_out()
        return self.create_response(message, {})
    
    async def handle_pause(self, message: Dict[str, Any]) -> Dict[str, Any]:
        """Handle pause request."""
        session = list(self.sessions.values())[0] if self.sessions else None
        if session:
            await session.pause()
        return self.create_response(message, {})
    
    async def handle_stack_trace(self, message: Dict[str, Any]) -> Dict[str, Any]:
        """Handle stackTrace request."""
        session = list(self.sessions.values())[0] if self.sessions else None
        if session:
            frames = await session.get_stack_trace()
            return self.create_response(message, {'stackFrames': frames, 'totalFrames': len(frames)})
        return self.create_response(message, {'stackFrames': [], 'totalFrames': 0})
    
    async def handle_scopes(self, message: Dict[str, Any]) -> Dict[str, Any]:
        """Handle scopes request."""
        args = message.get('arguments', {})
        frame_id = args.get('frameId')
        
        session = list(self.sessions.values())[0] if self.sessions else None
        if session:
            scopes = await session.get_scopes(frame_id)
            return self.create_response(message, {'scopes': scopes})
        return self.create_response(message, {'scopes': []})
    
    async def handle_variables(self, message: Dict[str, Any]) -> Dict[str, Any]:
        """Handle variables request."""
        args = message.get('arguments', {})
        variables_ref = args.get('variablesReference')
        
        session = list(self.sessions.values())[0] if self.sessions else None
        if session:
            variables = await session.get_variables(variables_ref)
            return self.create_response(message, {'variables': variables})
        return self.create_response(message, {'variables': []})
    
    async def handle_evaluate(self, message: Dict[str, Any]) -> Dict[str, Any]:
        """Handle evaluate request."""
        args = message.get('arguments', {})
        expression = args.get('expression')
        
        session = list(self.sessions.values())[0] if self.sessions else None
        if session:
            result = await session.evaluate(expression)
            return self.create_response(message, result)
        return self.create_error_response(message, "No active session")
    
    async def handle_disconnect(self, message: Dict[str, Any]) -> Dict[str, Any]:
        """Handle disconnect request."""
        for session in self.sessions.values():
            await session.disconnect()
        self.sessions.clear()
        return self.create_response(message, {})
    
    def create_response(self, request: Dict[str, Any], 
                       body: Dict[str, Any]) -> Dict[str, Any]:
        """Create success response."""
        return {
            'type': 'response',
            'request_seq': request.get('seq'),
            'success': True,
            'command': request.get('command'),
            'body': body,
            'seq': self.sequence,
        }
    
    def create_error_response(self, request: Dict[str, Any], 
                            error_message: str) -> Dict[str, Any]:
        """Create error response."""
        return {
            'type': 'response',
            'request_seq': request.get('seq'),
            'success': False,
            'command': request.get('command'),
            'message': error_message,
            'seq': self.sequence,
        }
    
    async def send_event(self, event: str, body: Dict[str, Any]):
        """Send DAP event."""
        message = {
            'type': 'event',
            'event': event,
            'body': body,
            'seq': self.sequence,
        }
        self.sequence += 1
        await self.send_message(message)


def main():
    """Entry point for debug adapter."""
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    adapter = VisioDebugAdapter()
    asyncio.run(adapter.start_server())


if __name__ == '__main__':
    main()
