# ADR-0002: Deterministic Layout & Pagination Engine

## Status
Accepted

## Context
Microsoft Word does not store discrete "page" elements in its OOXML file structure; pages are dynamically computed at runtime by Word's DirectWrite layout engine based on font metrics, paragraph spacing, line spacing, table cell dimensions, and margin bounds.
To provide accurate page counts, section header/footer isolation (such as cover pages with no headers/footers), and outline hierarchies in headless AI agent environments, `docx-agent` requires a deterministic server-side pagination engine.

## Decision
We implemented `LayoutEngine` (`src/docx_agent/engine/layout.py`), which:
1. Translates physical page dimensions (A4: 210mm x 297mm) and margins (Top 2.0cm, Bottom 2.0cm, Left 3.0cm, Right 2.0cm) into typographic points (pt).
2. Computes line-height and character wrapping for paragraphs (Times New Roman 13pt / 14pt / 16pt) accounting for Vietnamese unicode glyphs.
3. Computes exact row heights for tables by analyzing multi-line cell content, monospaced SQL code blocks (Consolas 10pt), cell padding, and borders.
4. Enforces `keep_with_next` orphan heading prevention and explicit page breaks (`pageBreakBefore`, `w:br[type='page']`).
5. Generates a `LayoutDocument` containing discrete `LayoutPage` objects, dynamic page counts (`Trang X / Y`), and a 3-level heading navigation outline.

## Consequences
### Positive:
- Server-side layout calculation achieves 1:1 page count fidelity with Microsoft Word (e.g. 66-67 pages on complex 70-question documents).
- Eliminates page count guesswork for automated document review and validation.

### Trade-offs:
- Advanced typographical features like multi-column complex hyphenation or curved WordArt require approximate bounding boxes.
