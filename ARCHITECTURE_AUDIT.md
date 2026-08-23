# ARCHITECTURE AUDIT: DOCX AGENT PLATFORM

**Date:** 2026-08-23  
**Status:** Audit Completed — Phase 0  
**Target:** Open-Source Universal DOCX Agent Platform (`docx-agent`)  
**Repository:** `https://github.com/TranNhatThach/Docx-Agent.git`

---

## 1. Executive Summary

This document performs an exhaustive audit of the existing DOCX tooling (`docx_tool.py` and associated skill metadata) and establishes the architectural specification, risk analysis, gap assessment, and migration strategy required to transform the codebase into a production-grade, agent-native **Universal DOCX Agent Platform**.

The platform is designed to be completely **agent-agnostic** (compatible with Antigravity, Claude Code, Cursor, Codex, Cline, Roo Code, Gemini, OpenAI, etc.) and operates under the core principle:

$$\text{LLM Proposes} \longrightarrow \text{Resolver Validates} \longrightarrow \text{Engine Executes (Atomic)} \longrightarrow \text{Verifier Confirms} \longrightarrow \text{Agent Reports}$$

---

## 2. Current Architecture & Codebase Analysis

The existing codebase consists of a single monolithic procedural script (`docx_tool.py`, ~209 lines) and an IDE-specific skill definition (`SKILL.md`).

```
Legacy Tooling
├── SKILL.md                 (Antigravity IDE specific skill manifest)
└── scripts/
    └── docx_tool.py         (Procedural argparse CLI wrapper around python-docx)
```

### Existing Capabilities Analyzed:
1. **`inspect` (`cmd_inspect`)**:
   - Prints paragraph count, table count, section count, and page geometry/margins for Section 0 only.
   - Iterates paragraphs looking for styles named `Heading*` to print a text-based outline.
2. **`read` (`cmd_read`)**:
   - Outputs paragraphs formatted with numerical prefixes `[P0000]`.
   - Optionally formats table rows with pipe `|` characters.
3. **`replace` (`cmd_replace`)**:
   - Checks if target substring exists in a single run; if so, replaces `r.text`.
   - If substring spans across multiple runs, falls back to `p.text = p.text.replace(...)`.
   - In tables, performs `p.text = p.text.replace(...)` unconditionally on matching paragraphs.
4. **`edit-para` (`cmd_edit_para`)**:
   - Looks up paragraph by 0-based array index and assigns `p.text = args.text`.
5. **`append-para` (`cmd_append_para`)**:
   - Calls `doc.add_paragraph()` or `doc.add_heading()`.
6. **`md2docx`**:
   - Documented in comments/docstrings but completely unimplemented in code (stub/ghost feature).

---

## 3. Critical Problems & Vulnerabilities Identified

### 3.1 Catastrophic Formatting Destruction
- **The `p.text = ...` Anti-pattern**: In `docx-python`, assigning to `paragraph.text` clears the entire paragraph XML child element list (`w:p`) and inserts a single plain text run (`w:r` -> `w:t`).
- **Impact**: All inline formatting is permanently lost:
  - Font families, sizes, colors, highlights
  - Bold, italic, underline, strikethrough, superscript, subscript
  - Hyperlinks (`w:hyperlink`), bookmarked ranges, inline drawings/images
  - Run-level properties and character styles
  - Field codes (`w:fldSimple`, `w:instrText`)
- **Cross-Run Replacement Failure**: In `cmd_replace`, whenever a search term crosses run boundaries (e.g. `[Run 1: "Hello "] [Run 2 (Bold): "World"]` matching `"Hello World"`), the script triggers `p.text = ...`, wiping out the bold styling of `"World"`. In tables, it executes this destructive fallback immediately.

### 3.2 Fragile Index-Based Targeting & Lack of Stable Identity
- Operations depend entirely on transient integer indices (`--index 15`).
- Inserting or deleting a single paragraph invalidates every subsequent index.
- There is no stable element UUID / identity scheme for paragraphs, tables, rows, cells, headers, or footers.

### 3.3 Zero Transaction Safety & Atomicity
- In-place modifications directly overwrite the source file via `doc.save(path)`.
- If an unhandled exception occurs during execution, the target document is left in a corrupted or half-modified state.
- No backup mechanism (`.bak` or timestamped shadow copies).
- No staging sandbox (`temp_file -> reopen -> verify -> atomic rename/commit`).

### 3.4 No Real Verification ("Blind Success")
- Current code assumes that if `doc.save()` completes without raising a Python exception, the operation succeeded.
- It never validates XML schemas, structural integrity, relationship preservation, or whether formatting matches the intended target state.

### 3.5 Missing Subsystems
- **No OOXML Field Engine**: Cannot generate true dynamic page numbers (`PAGE`, `NUMPAGES`), document titles (`TITLE`), dates (`DATE`), or Tables of Contents (`TOC`).
- **No Style Engine**: Cannot inspect, create, update, or inherit paragraph/character styles programmatically.
- **No Layout & Section Engine**: Multi-section documents, header/footer unlinking, column layouts, and custom margins are unsupported.
- **No Academic & Preset System**: No support for formal document standards (e.g. Vietnamese Academic `academic-vn`, IEEE, APA, Technical Reports).
- **No Diff Engine**: Agents cannot compute structured semantic diffs between document revisions.
- **No Plan / Batch Execution**: Agents cannot submit transactional change manifests with dry-run conflict detection.
- **No Machine-Readable Standard**: Stdout contains arbitrary human-readable print statements instead of structured, schema-validated JSON outputs.

---

## 4. Target Architecture & Component Design

The new platform is structured around strict separation of concerns, transaction-safety, format preservation, and multi-interface support (CLI, MCP, Python API).

```
                      AI CODING AGENTS
            (Antigravity / Claude Code / Cursor / Codex / Cline / ...)
                               │
            ┌──────────────────┼──────────────────┐
            ▼                  ▼                  ▼
     CLI Interface        MCP Server         Python API
   (docx-agent CLI)   (JSON-RPC Tooling)  (DocumentAgent)
            │                  │                  │
            └──────────────────┼──────────────────┘
                               │
                     ┌─────────▼─────────┐
                     │   Protocol & DTO  │
                     │  (Pydantic Models)│
                     └─────────┬─────────┘
                               │
                     ┌─────────▼─────────┐
                     │   DocumentAgent   │  <--- Orchestrator
                     └─────────┬─────────┘
                               │
         ┌─────────────────────┼─────────────────────┐
         ▼                     ▼                     ▼
┌─────────────────┐   ┌─────────────────┐   ┌─────────────────┐
│ Core & Resolver │   │ Operations      │   │ Transactions    │
│ - DocumentModel │   │ - ContentOps    │   │ - BackupManager │
│ - ElementID     │   │ - FormattingOps │   │ - RollbackGuard │
│ - TargetResolver│   │ - TableOps      │   │ - AtomicCommit  │
│ - CapabilityRep │   │ - StyleOps      │   │ - StagingTemp   │
└─────────────────┘   │ - SectionOps    │   └─────────────────┘
                      │ - ImageOps      │
                      │ - AcademicOps   │
                      └────────┬────────┘
                               │
         ┌─────────────────────┼─────────────────────┐
         ▼                     ▼                     ▼
┌─────────────────┐   ┌─────────────────┐   ┌─────────────────┐
│ OOXML Low-Level │   │ Verification    │   │ Diff & Presets  │
│ - Runs & Split  │   │ - DocxValidator │   │ - SemanticDiff  │
│ - Fields (PAGE) │   │ - FormatChecker │   │ - PresetLoader  │
│ - NumberingDef  │   │ - IntegrityReopen│  │ - Academic VN   │
│ - TOC Generator │   │ - LayoutChecker │   │ - IEEE / APA    │
└─────────────────┘   └─────────────────┘   └─────────────────┘
                               │
                               ▼
                       python-docx + lxml
                               │
                               ▼
                      OpenXML (.docx) Files
```

---

## 5. Subsystem Specifications

### 5.1 Document Model & Stable Identity System
- Every document element is assigned a stable deterministic context ID:
  - Paragraphs: `p_0001`, `p_0002`...
  - Tables: `tbl_0001`, `tbl_0002`...
  - Rows: `tbl_0001_r01`...
  - Cells: `tbl_0001_r01_c01`...
  - Sections: `sec_0001`...
  - Runs: `p_0001_r01`...
- **Target Resolver**: Resolves targets flexibly via:
  - Element ID (`p_0012`, `tbl_0002`)
  - Paragraph index (`idx:12`)
  - Exact text match (`text:"Phương pháp nghiên cứu"`)
  - Regular expressions (`regex:"^Chương \\d+"`)
  - Heading hierarchy (`heading:level=2,text="Kết quả"`)
  - Style selector (`style:"Caption"`)
  - Semantic range selector (`range:p_0005..p_0020`)

### 5.2 Format-Preserving Run Surgery Engine (`ooxml/runs.py`)
- Represents paragraph text as a character-to-run position map.
- When replacing a substring that spans multiple runs:
  1. Computes the exact slice intersection across runs.
  2. Splits boundary runs if the substring starts or ends mid-run, cloning all run formatting (`rPr`).
  3. Applies the replacement text directly to the target run segments while leaving unaffected text and formatting untouched.
  4. Preserves hyperlinks, embedded fields, bold, italic, color, font size, and language settings.

### 5.3 Style & Academic Presets Engine
- Built-in JSON presets (`academic_vn.json`, `ieee.json`, `apa.json`, `technical_report.json`).
- `academic_vn` specification:
  - Page: A4 ($21.0 \times 29.7\text{ cm}$)
  - Margins: Top $2.0\text{ cm}$, Bottom $2.0\text{ cm}$, Left $3.0\text{ cm}$, Right $2.0\text{ cm}$
  - Body: Times New Roman, $13\text{ pt}$, 1.5 line spacing, Justified, First-line indent $1.27\text{ cm}$ ($0.5\text{ in}$)
  - Headings: Standard hierarchical sizing ($16\text{ pt}$ Bold, $14\text{ pt}$ Bold, $13\text{ pt}$ Bold/Italic), Keep-with-next enabled.

### 5.4 Transaction & Verification Pipeline
- Execution lifecycle for every mutation:
  ```
  VALIDATE -> BACKUP -> LOAD -> PLAN -> APPLY -> SAVE_TEMP -> REOPEN_INDEPENDENT -> VERIFY -> COMMIT
  ```
- If verification fails at any check:
  - Staging temp file is discarded.
  - Rollback is triggered.
  - Detailed diagnostic payload with actual vs. expected discrepancies is returned to the agent for self-repair.

### 5.5 Multi-Interface Architecture
1. **MCP Server**: Standard Model Context Protocol server exposing high-level tools (`docx_inspect`, `docx_read`, `docx_find`, `docx_replace`, `docx_format_text`, `docx_format_paragraph`, `docx_style`, `docx_table`, `docx_image`, `docx_section`, `docx_header`, `docx_footer`, `docx_toc`, `docx_verify`, `docx_diff`, `docx_apply_plan`).
2. **CLI**: Comprehensive rich command-line tool with `--json`, `--verbose`, and backward-compatible alias wrappers.
3. **Python SDK**: Clean, object-oriented `DocumentAgent` class for direct programmatic consumption.

---

## 6. Refactoring & Migration Strategy

1. **Maintain Zero-Breakage Compatibility**:
   - The legacy CLI commands (`inspect`, `read`, `replace`, `edit-para`, `append-para`, `md2docx`) will be supported as compatibility adapters delegating directly to `DocumentAgent`.
2. **Package Structure**:
   - Standard PEP 621 compliant package with `pyproject.toml`.
   - Entry points: `docx-agent` (CLI) and `docx-agent-mcp` (MCP server).
3. **Phased Implementation & Quality Gates**:
   - 15 incremental phases with strict test coverage, real DOCX fixture generation, and independent verification.
