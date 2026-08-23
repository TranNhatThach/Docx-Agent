# Contributing to Docx-Agent Platform

Thank you for contributing to the **Universal Open-Source DOCX Agent Platform & AI-Native Document Workspace**!

---

## 1. Quickstart Development Setup

```bash
# 1. Clone your fork
git clone https://github.com/your-username/docx-agent.git
cd docx-agent

# 2. Create virtual environment
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate

# 3. Install in editable mode with development & MCP dependencies
pip install -e ".[dev,mcp]"

# 4. Install pre-commit hooks
pre-commit install
```

---

## 2. Developer Workflow Commands (Makefile)

We provide a standard `Makefile` for developer tasks:

```bash
make test        # Run entire pytest suite
make test-cov    # Run test suite with HTML & terminal coverage report
make lint        # Lint codebase using Ruff
make format      # Auto-format codebase using Ruff
make typecheck   # Static type check using Mypy
make build       # Build distribution packages (wheel + sdist)
make clean       # Clean caches and build artifacts
```

---

## 3. Git Branching & Commit Conventions

### Branch Naming Conventions
- `feature/<short-name>`: New capabilities or adapters.
- `fix/<short-name>`: Bug fixes and layout corrections.
- `refactor/<short-name>`: Code restructuring without feature change.
- `docs/<short-name>`: Documentation and ADR updates.
- `security/<short-name>`: Security vulnerabilities and dependency patches.

### Conventional Commits
All commits must follow the [Conventional Commits](https://www.conventionalcommits.org/) standard:
- `feat: add DrawingML image alignment support`
- `fix: correct table cell height calculation in layout engine`
- `refactor: extract style cascade into StyleResolver`
- `docs: add ADR-0003 dual rendering strategy`
- `test: add 50-page layout stress regression test`

---

## 4. Core Engineering Rules

1. **Format Preservation is Paramount**:
   - Never overwrite paragraph text with `p.text = ...` directly.
   - Always perform run surgeries preserving `<w:rPr>` properties.
2. **Deterministic Element Identities**:
   - Target operations through `TargetResolver` and `IdentityManager`.
3. **Transaction Safety**:
   - Mutations must operate within `TransactionContext` with shadow backup and staging sandboxes.
4. **Verification-First**:
   - Verify every mutation with `DocumentValidator` before committing changes.
5. **No Regressions**:
   - Every fixed bug must be backed by a test in `tests/regression/` or `tests/rendering/`.
6. **Multi-Interface Neutrality**:
   - Core domain logic lives in `canonical/` and `engine/`. CLI and MCP servers must remain thin interface adapters.

---

## 5. Pull Request Checklist

Before submitting a Pull Request:
- [ ] All tests pass (`make test`).
- [ ] Code is formatted and linted without errors (`make lint`).
- [ ] Type checks pass (`make typecheck`).
- [ ] Any new feature is documented in `docs/` and `CHANGELOG.md`.
- [ ] PR description follows `.github/PULL_REQUEST_TEMPLATE.md`.
