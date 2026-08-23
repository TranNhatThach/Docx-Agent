# Architecture Dependency Rules & Boundaries

To ensure Docx-Agent remains maintainable, testable, and resistant to architectural decay over years of contributions, the following dependency direction rules are strictly enforced:

## 1. Direction of Dependencies
```text
Interfaces  ───>  Application  ───>  Domain / Core
      │                 │                  ▲
      │                 │                  │
      └──────────> Engine / Adapters ──────┘
```

- **Domain/Core** (`canonical/`, `core/`): Must **NEVER** depend on UI frameworks, CLI libraries, HTTP servers, or external network adapters.
- **Engine** (`engine/`): Depends only on `canonical/` and `core/`. Must remain headless and runnable in CLI/tests without a browser or display server.
- **Adapters** (`adapters/`): Handle the translation between external storage/formats (OpenXML, Markdown) and the canonical domain model.
- **Interfaces** (`interfaces/`): Depend on Application and Domain layers. Webviews and CLIs consume domain payloads and must not embed core business logic.

## 2. Prohibition of Circular Dependencies
- No module inside `docx_agent.canonical` may import from `docx_agent.engine` or `docx_agent.adapters`.
- No module inside `docx_agent.core` may import from `docx_agent.interfaces`.
- Use dependency injection or callback interfaces when the engine needs to notify upper layers.

## 3. Public API vs Internal Implementation
- Only symbols explicitly exported in `docx_agent.__all__` are considered stable Public APIs.
- Submodules prefixed with `_` or internal helper functions in `docx_agent.ooxml` are private and subject to refactoring across minor releases.
