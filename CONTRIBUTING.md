# Contributing to Docx-Agent Platform

Thank you for contributing to the **Universal Open-Source DOCX Agent Platform**!

## Development Guidelines

1. **Format Preservation is Paramount**:
   - Never use `p.text = ...` directly in paragraph mutations.
   - Always perform run surgeries and maintain `<w:rPr>` properties.
2. **Deterministic Element Identities**:
   - Operations must resolve through `TargetResolver` and `IdentityManager`.
3. **Transaction Safety**:
   - Mutations must go through `TransactionContext` with shadow backup and staging sandboxes.
4. **Verification-First**:
   - Never claim success without independent reopen and verification checks.
5. **Multi-Interface Neutrality**:
   - Keep business logic in `DocumentAgent` and core modules. CLI and MCP servers must remain thin adapters.

## Running Tests

```bash
pip install -e ".[dev,mcp]"
pytest tests/ -v
```

## Pull Request Process

1. Fork the repository and create a feature branch.
2. Ensure all unit, regression, and E2E tests pass.
3. Submit PR with detailed description and test verification logs.
