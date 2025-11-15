# Changelog

All notable changes to visiowings will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Remote VBA debugging support via Debug Adapter Protocol (DAP)
- VS Code integration for debugging Visio VBA code
- Debug adapter server with full DAP implementation
- Thread-safe COM bridge for Visio automation
- Breakpoint injection and management system
- Debug session coordination with reconnect support
- CLI commands for debug server management (`visiowings-debug`)
- Comprehensive debugging documentation and examples
- Unit tests for debugging components
- Example VBA modules for testing debugger

### Features
- Set and remove breakpoints in VBA code
- Step execution (over, in, out)
- Call stack navigation
- Continue and pause execution
- Session resilience with reconnection
- Multiple concurrent debug sessions
- Breakpoint verification and restoration
- Error handling with automatic cleanup

### Technical Details
- Implements Debug Adapter Protocol (DAP) for VS Code
- Uses pywin32 for COM automation with Visio VBA
- Asyncio-based architecture for non-blocking operations
- Thread-safe COM operations with dedicated worker thread
- Breakpoint injection using VBA Stop statements
- Original code preservation and restoration
- Configurable timeouts and retry logic

## [0.2.0] - Previous Release

### Added
- Initial VBA export/import functionality
- File watching with automatic synchronization
- OneDrive path resolution
- VBA header auto-fix
- COM initialization improvements

### Fixed
- Empty module deletion issues
- COM threading initialization
- Error handling improvements
- Path matching for cloud storage

## [0.1.0] - Initial Release

### Added
- Basic VBA code export from Visio documents
- Basic VBA code import to Visio documents
- Command-line interface
- Core document management
- Visio COM connection handling

[Unreleased]: https://github.com/twobeass/visiowings/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/twobeass/visiowings/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/twobeass/visiowings/releases/tag/v0.1.0
