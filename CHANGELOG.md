# Changelog

All notable changes to the `docx-agent` platform will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-08-23

### Added
- **Core Architecture**: `DocumentAgent` orchestrator, `DocumentModel`, and polymorphic `TargetResolver`.
- **Run Surgery Engine**: Format-preserving cross-run text replacement engine with `<w:rPr>` cloning.
- **Stable Identity System**: Deterministic context IDs (`p_0001`, `tbl_0001`, `sec_0001`).
- **Formatting Engine**: Font family, font size, bold, italic, underline, strike, colors, alignments, line spacing, margins, indents.
- **Style Engine**: Inspection, creation, updating, and applying paragraph/character styles.
- **Academic Presets**: Built-in institutional presets (`academic-vn`, `ieee`, `apa`, `technical-report`).
- **Tables & Images**: Table creation, cell editing, border styling, header row repeating, and image insertion with captions.
- **OOXML Fields & TOC**: Native Word dynamic fields (`PAGE`, `NUMPAGES`, `DATE`, `TOC`).
- **Transaction & Verification**: Atomic staging, pre-mutation shadow backup, independent reload validator, and format checking.
- **Semantic Diff Engine**: Comprehensive document revision diffing returning structured JSON.
- **Interfaces**: Rich Typer CLI with `--json` support, stdio MCP server (`docx-agent-mcp`), and Python SDK.
- **Markdown Converter**: `md2docx` high-fidelity renderer.
- **Test Suite**: 26 unit, integration, large-document stress, and 15-step master E2E acceptance tests.
