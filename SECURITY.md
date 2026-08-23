# Security Policy

## 1. Supported Versions

We provide security updates and patches for the following versions:

| Version | Supported          | Security Patches |
| ------- | ------------------ | ---------------- |
| 2.1.x   | :white_check_mark: | Full Active Support |
| 2.0.x   | :white_check_mark: | Critical Fixes Only |
| 1.0.x   | :x:                | Deprecated |

---

## 2. Security Architecture Principles

`docx-agent` is engineered for autonomous AI agents and enterprise document pipelines with strict defense-in-depth:

1. **No Arbitrary Code Execution**: The platform strictly refuses execution of embedded Word VBA macros (`.docm`), external OLE binaries, or script attachments.
2. **Path Traversal & Sandboxing**: All file paths are strictly resolved through `src/docx_agent/utils/paths.py` with boundary checks, preventing arbitrary file write outside designated document directories.
3. **Zip Bomb & Memory Defense**: OpenXML packages are subject to size limits (`max_file_size_mb = 50`) and XML entity expansion protection to prevent resource exhaustion attacks.
4. **Log Sanitization**: Structured logs scrub potential secrets, authorization tokens, and credentials.
5. **Atomic File Rollback**: Destructive operations cannot leave corrupted partial writes on disk due to automatic `.bak` transaction rollback.

---

## 3. Reporting a Vulnerability

If you discover a security vulnerability within `docx-agent`:
- **Please DO NOT open a public issue.**
- Submit a private report via **GitHub Security Advisories** or email the maintainers directly at `thachtn@example.com`.
- Include:
  - Description of the vulnerability.
  - Step-by-step reproduction instructions or proof-of-concept `.docx`.
  - Potential impact and affected versions.

We commit to acknowledging your report within **48 hours** and providing a mitigation timeline.
