# ADR-0001: Canonical Document Model as Single Source of Truth

## Status
Accepted

## Context
Traditional DOCX automation libraries (such as raw `python-docx` or basic HTML conversion bridges) either:
1. Couple the application directly to complex, mutable XML DOM trees with low abstraction.
2. Convert DOCX to rudimentary HTML/Markdown, discarding sections, multilevel numbering, header/footer isolation, and table grid coordinates.

AI coding agents and visual workspace editors need a clean, deterministic, typed in-memory model that preserves 100% of WordprocessingML semantics while providing fast query, diff, and mutation APIs.

## Decision
We established a typed, Pydantic-based **Canonical Document Model** (`DocumentNode`, `SectionNode`, `ParagraphBlock`, `HeadingBlock`, `TableBlock`, `ListItemBlock`, `ImageBlock`, `DiagramBlock`, `UnsupportedBlock`) as the single authoritative in-memory representation.
- Every node has a collision-resistant UUID (`id`).
- Direct formatting and cascaded effective formatting are cleanly separated.
- Unsupported/foreign OOXML elements are captured in `UnsupportedBlock` and roundtripped safely without data loss.

## Consequences
### Positive:
- Total decoupling of application/agent logic from the underlying storage format.
- Seamless JSON serialization for webviews, MCP servers, and LLM tool calling.
- Enables immutable diffing, provenance tracking (`HUMAN`, `AGENT`, `IMPORT`), and transactional rollbacks.

### Trade-offs:
- Requires dedicated `DocxImporter` and `DocxExporter` adapters to synchronize between OOXML and the Canonical Model.
