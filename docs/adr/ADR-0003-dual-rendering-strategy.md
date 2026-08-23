# ADR-0003: Dual-Engine Rendering Strategy (DOCX-Native Binary + Canonical RenderTree)

## Status
Accepted

## Context
Visual document editing in IDE webviews requires both:
1. Pixel-perfect visual fidelity matching Microsoft Word (including exact borders, shadings, headers, footers, and page numbers).
2. Deep agentic integration where an AI coding agent can inspect structural nodes, modify AST blocks, and preview semantic diffs.

## Decision
We adopted a **Dual-Engine Rendering Strategy**:
1. **Primary UI Engine (OpenXML Binary Native)**: The VS Code extension host streams the raw `.docx` binary buffer (Base64) to the Webview. The Webview parses the binary package via `docx-preview` / `JSZip`, rendering genuine WordprocessingML pages into the DOM with in-place live editing (`contentEditable`).
2. **Deterministic Agent Engine (Python Canonical RenderTree)**: Serves as the fallback in offline/standalone environments and acts as the semantic engine for AI agent selection contexts, AST operations, and programmatic inspection.

## Consequences
### Positive:
- The UI achieves 100% visual fidelity matching Microsoft Word.
- Seamless fallback ensures zero breakage in offline or headless testing environments.
- AI agents operate on high-level semantic nodes rather than fragile browser DOM selections.

### Trade-offs:
- Webview bundle includes `docx-preview` and `JSZip` client scripts.
