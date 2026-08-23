# Universal DOCX Agent Platform (`docx-agent`)

[![CI Tests](https://img.shields.io/badge/tests-26%20passed-brightgreen.svg)]()
[![Python Version](https://img.shields.io/badge/python-3.10%2B-blue.svg)]()
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![MCP Protocol](https://img.shields.io/badge/MCP-2024--11--05-purple.svg)]()

**A universal, agent-native Microsoft Word (`.docx`) manipulation engine that allows AI coding agents to read, understand, modify, format, generate, validate, and repair Word documents safely.**

---

## 🌟 Vision & Design Philosophy

Traditional script-based tools manipulate Word documents by overwriting paragraph text (`p.text = ...`), causing catastrophic destruction of inline run properties, bold/italic stylings, font families, colors, hyperlinks, and document structure.

`docx-agent` is built from the ground up on the **Agent-Native Document Protocol**:

$$\text{LLM Proposes} \longrightarrow \text{Resolver Validates} \longrightarrow \text{Engine Executes (Atomic)} \longrightarrow \text{Verifier Confirms} \longrightarrow \text{Agent Reports}$$

### Core Tenets:
- **Agent-Agnostic Core**: Antigravity, Claude Code, Cursor, Codex, Cline, and Roo Code connect seamlessly via standard MCP (Model Context Protocol) and CLI.
- **Strict Format Preservation**: Deep character-to-run surgery preserves unaffected formatting, colors, and hyperlinks even during multi-run substring replacements.
- **Atomic Transactions & Rollback**: Every mutation stages changes to a sandbox, reopens independently, verifies integrity, and rolls back cleanly if validation fails.
- **No Fake Features**: True OpenXML field integration for dynamic page numbering (`PAGE` / `NUMPAGES`), Tables of Contents (`TOC`), and table headers.
- **Institutional Presets**: Built-in styling engines for Vietnamese Academic theses (`academic-vn`), IEEE, APA, and Technical Reports.

---

## 🏗 Architecture

```
                  AI CODING AGENTS
  (Antigravity / Claude Code / Cursor / Codex / Cline / Roo Code)
                         │
       ┌─────────────────┼─────────────────┐
       │                 │                 │
      CLI               MCP           Python API
  (docx-agent)    (docx-agent-mcp)  (DocumentAgent)
       │                 │                 │
       └─────────────────┼─────────────────┘
                         │
              ┌──────────▼──────────┐
              │   DOCX AGENT CORE   │
              │                     │
              │ Document Model      │
              │ Target Resolver     │
              │ Run Surgery Engine  │
              │ Style & Presets     │
              │ Tables & Images     │
              │ Layout & Sections   │
              │ Independent Verifier│
              │ Atomic Transactions │
              └──────────┬──────────┘
                         │
                ┌────────▼────────┐
                │ python-docx     │
                │ + OOXML Engine  │
                └────────┬────────┘
                         │
                         ▼
                Microsoft Word (.docx)
```

---

## 🚀 Installation

Install directly via `pip` or `pipx`:

```bash
pip install docx-agent
```

For development and local testing:

```bash
git clone https://github.com/TranNhatThach/Docx-Agent.git
cd Docx-Agent
pip install -e ".[dev,mcp]"
```

---

## 💻 CLI Usage

The CLI supports both rich visual output for humans and machine-readable JSON for AI coding agents (`--json`).

### 1. Inspection & Discovery
```bash
# Inspect document structure, geometries, margins, and headings outline
docx-agent inspect report.docx

# Read paragraphs with deterministic element IDs (p_0001, p_0002...)
docx-agent read report.docx --start 0 --end 20 --json

# Find paragraphs matching substring, regex, or style
docx-agent find report.docx --text "Phương pháp" --json
docx-agent find report.docx --regex "^Chương \d+" --json

# Check document capabilities and unsupported element warnings
docx-agent capabilities report.docx --json
```

### 2. Format-Preserving Content Operations
```bash
# Replace text while strictly preserving surrounding bold/italic/color runs
docx-agent replace report.docx --target "thuật toán cũ" --replace "mô hình học sâu"

# Insert paragraph before/after target element
docx-agent insert report.docx --text "Nội dung bổ sung" --target "p_0012" --position "after"

# Append paragraph or heading
docx-agent append-para report.docx --text "Chương 4: Kết quả" --style "Heading 1"
```

### 3. Typography, Spacing & Presets
```bash
# Apply full Vietnamese Academic Thesis preset (A4, TNR 13pt, 1.5 spacing, justified)
docx-agent preset report.docx --name "academic-vn"

# Format character font properties
docx-agent format-text report.docx --target "p_0005" --font-name "Times New Roman" --font-size-pt 14 --bold

# Format paragraph layout and spacing
docx-agent format-paragraph report.docx --target "p_0005" --alignment "justify" --line-spacing 1.5 --first-line-indent-cm 1.27
```

### 4. Verification, Diff & Markdown Conversion
```bash
# Independently verify document integrity and typography conformance
docx-agent verify report.docx --expected-font "Times New Roman" --expected-size 13.0 --json

# Semantic diff between two DOCX revisions
docx-agent diff before.docx after.docx --json

# Convert Markdown file to professionally styled DOCX
docx-agent md2docx report.md --output report.docx --preset "academic-vn"
```

---

## 🔌 Model Context Protocol (MCP) Server

`docx-agent` provides a first-class MCP server exposing high-level, format-preserving tools to any MCP-compatible coding agent.

### Launch Command
```bash
docx-agent-mcp
```

### Configuration Examples

#### Antigravity / Cursor / Claude Code (`mcp_config.json`):
```json
{
  "mcpServers": {
    "docx-agent": {
      "command": "docx-agent-mcp",
      "args": []
    }
  }
}
```

### Core MCP Tools Exposed:
- `docx_inspect`: Structural summary and heading outline.
- `docx_read`: Numbered paragraphs with stable element IDs.
- `docx_find`: Search text, regex patterns, or styles.
- `docx_replace`: Surgical format-preserving text replacements.
- `docx_insert`: Element insertion before/after reference targets.
- `docx_format_text`: Font family, sizing, bold, italic, colors.
- `docx_format_paragraph`: Alignment, line spacing, margins.
- `docx_preset`: Apply presets (`academic-vn`, `ieee`, `apa`, `technical-report`).
- `docx_table`: Create and format data tables with repeating headers.
- `docx_image`: Insert, position, and caption images.
- `docx_verify`: Independent integrity and format verification.
- `docx_diff`: Semantic revision diffing.
- `docx_apply_plan`: Atomic batch plan execution with rollback.

---

## 🐍 Python SDK API

```python
from docx_agent import DocumentAgent

# Load or create document
agent = DocumentAgent("thesis.docx")

# Inspect document outline
outline = agent.outline()

# Apply institutional thesis preset
agent.apply_preset("academic-vn")

# Replace text without breaking inline bold / italic stylings
agent.replace("mô hình BERT", "mô hình RoBERTa")

# Insert data table
table_id = agent.create_table(
    rows=3,
    cols=2,
    data=[["Tham số", "Giá trị"], ["Epochs", "100"], ["Accuracy", "99.2%"]],
)

# Insert Table of Contents
agent.insert_toc(title="MỤC LỤC")

# Save with automatic transaction sandboxing and independent verification
agent.save("thesis_final.docx", verify=True)
```

---

## 🛡 Transaction & Verification Safety Model

Every mutation executed by `docx-agent` follows the strict atomic lifecycle:

```
VALIDATE SOURCE
      ↓
SHADOW BACKUP (.bak)
      ↓
LOAD & PARSE
      ↓
EXECUTE RUN SURGERY
      ↓
SAVE TO TEMP STAGING
      ↓
INDEPENDENT REOPEN & VERIFY
      ↓
COMMIT / ROLLBACK ON FAILURE
```

---

## 🧪 Testing

The platform includes exhaustive unit, regression, and end-to-end integration tests:

```bash
pytest tests/ -v
```

- **Format Preservation Tests**: Verifies run boundary preservation across single and multi-run replacements.
- **Vietnamese Unicode Fidelity**: Tests 100% accurate rendering of Vietnamese accents, typography marks, and mathematical symbols.
- **Large Document Stress Tests**: Benchmarks execution against 1000+ paragraphs and 50+ tables.
- **15-Step Master E2E Workflow**: End-to-end simulation covering presets, edits, tables, images, TOC, verification, and diffing.

---

## 📄 License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.
