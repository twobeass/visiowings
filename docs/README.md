# Visiowings Implementation Documentation

This directory contains comprehensive technical documentation for the visiowings project.

## Documentation Index

### Architecture & Design
- [Architecture Overview](architecture.md) - System design, component interaction, and data flow
- [Design Principles](design-principles.md) - Core principles guiding the implementation

### Component Documentation
- [CLI Module](components/cli.md) - Command-line interface and argument parsing
- [Document Manager](components/document-manager.md) - Multi-document VBA project management
- [VBA Export](components/vba-export.md) - Export logic, header stripping, and file comparison
- [VBA Import](components/vba-import.md) - Import logic, header repair, and encoding handling
- [File Watcher](components/file-watcher.md) - Bidirectional sync and change detection
- [Visio Connection](components/visio-connection.md) - COM automation interface

### Technical Specifications
- [File Formats](specs/file-formats.md) - VBA module file structure and header formats
- [Encoding Handling](specs/encoding.md) - Character encoding strategy (cp1252 ↔ UTF-8)
- [Change Detection](specs/change-detection.md) - Hash-based and content comparison algorithms

### Developer Guides
- [Development Setup](dev/setup.md) - Environment setup and dependencies
- [Testing Guide](dev/testing.md) - Testing strategy and guidelines
- [Contributing](dev/contributing.md) - Contribution guidelines and code standards
- [Troubleshooting Development](dev/troubleshooting.md) - Common development issues

## Quick Navigation

**New to the codebase?** Start with [Architecture Overview](architecture.md)

**Want to contribute?** Read [Contributing Guide](dev/contributing.md)

**Debugging an issue?** Check [Troubleshooting](dev/troubleshooting.md)

**Understanding file formats?** See [File Formats](specs/file-formats.md)
