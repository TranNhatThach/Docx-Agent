# Security Policy

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| 1.0.x   | :white_check_mark: |

## Security Considerations for AI Coding Agents

`docx-agent` is designed with defense-in-depth principles for autonomous AI agents:

1. **No Arbitrary Code Execution**: The platform strictly refuses execution of embedded Word VBA macros or external executable scripts.
2. **Safe Path Resolution**: File paths are normalized and resolved safely to prevent path traversal outside designated workspace directories.
3. **Atomic Rollback**: Destructive operations cannot corrupt user files without automatic restoration.

## Reporting a Vulnerability

If you discover a security vulnerability within `docx-agent`, please create a private issue on GitHub or contact the maintainer directly at `thachtn@example.com`.
