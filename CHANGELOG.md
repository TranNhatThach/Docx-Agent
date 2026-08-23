# Changelog

All notable changes to the **Docx-Agent** platform will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [2.0.0] - 2026-08-23

### Major Architecture Upgrade: AI-Native Document Workspace
- **Canonical Document Model**: Decoupled in-memory runtime representation (`DocumentNode`, `SectionNode`, `BlockNode`, `RunNode`, `CitationNode`, `DiagramBlock`, `UnsupportedBlock`).
- **Deterministic Operation Engine**: Reversible command architecture (`InsertTextOp`, `DeleteTextOp`, `ReplaceTextOp`, `FormatParagraphOp`, `InsertBlockOp`, `InsertCitationOp`, `CompositeOperation`).
- **Selection-Aware Collaboration**: Rich selection context extractor capturing character ranges, surrounding paragraphs, active section headers, and document profiles.
- **Agent Transaction Manager**: Multi-operation transaction staging, live preview diffs, atomic commit, and single-click full transaction rollback.
- **Research Assistant & Citation Engine**: Real peer-reviewed evidence matching (APA, IEEE, Academic-VN) with strict anti-hallucination verification.
- **Diagram & Media Synthesizer**: Native Mermaid/SVG generation for system architecture, flowcharts, sequence diagrams, and use cases.
- **Ambiguity & Clarification Engine**: Confidence assessment triggering structured multiple-choice clarification questions on underspecified prompts.
- **Dual Verification Engine**: Structural XML integrity + visual layout verification (image/table overflows, orphan headings, excessive whitespace).
- **Interactive Visual Workspace**: Web-based A4 document canvas with typography toolbar, outline navigator, and live Agent chat.
- **Full Backward Compatibility**: 100% test pass rate across all legacy v1.0 CLI commands, MCP tools, and Python SDK methods.

---

## [1.0.0] - 2026-08-23

### Added
- Core Open-Source DOCX Agent Platform.
- Run Surgery Engine for format-preserving cross-run text replacement.
- Deterministic element identity mapping (`p_0001`, `tbl_0001`, `sec_0001`).
- Typography and style engine (font families, sizes, colors, line spacing, margins, indents).
- Academic thesis presets (`academic-vn`, `ieee`, `apa`, `technical-report`).
- Table generator with repeating headers (`w:tblHeader`) and cell shading.
- Dynamic OOXML fields (`PAGE`, `NUMPAGES`, `DATE`, `TOC`).
- Atomic transactions with shadow `.bak` backups and independent reopen verification.
- Model Context Protocol (MCP) server on stdio.
- Typer CLI with `--json` machine-readable output.
- Markdown to DOCX converter (`md2docx`).
- 26 initial test cases.
