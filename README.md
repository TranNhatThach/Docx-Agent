# Docx-Agent V2: AI-Native Document Workspace & Engine 🚀

[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Tests](https://img.shields.io/badge/tests-36%2F36%20passing-brightgreen.svg)]()
[![Architecture: V2 Workspace](https://img.shields.io/badge/Architecture-AI--Native%20Workspace-orange.svg)]()

> **Universal, Agent-Native Microsoft Word (`.docx`) Manipulation Engine & Visual Document Workspace for Humans & AI Coding Agents.**

---

## 🌟 What is Docx-Agent V2?

Docx-Agent V2 transforms traditional file-level Word editing into a **unified AI-native workspace**. It bridges human writers, visual document canvases, and AI coding agents (Antigravity, Claude Code, Cursor, Codex, Cline, Roo Code) through a high-performance **Canonical Document Model** and deterministic **Operation Engine**.

```
    HUMAN                         AI CODING AGENTS
      ↕                                  ↕
VISUAL DOCUMENT WORKSPACE  ↔  SELECTION CONTEXT & AGENT CHAT
      ↕                                  ↕
            CANONICAL DOCUMENT ENGINE & TX MANAGER
                                 ↕
            DOCX IMPORT / EXPORT / DUAL VERIFICATION
                                 ↕
                       Microsoft Word (.docx)
```

---

## 🚀 Key Features

### 1. Dual Representation & Decoupled State
- **Canonical Document Model**: Documents are represented in-memory as a high-performance node tree (`DocumentNode`, `SectionNode`, `BlockNode`, `RunNode`, `CitationNode`).
- **DOCX as Interchange Format**: Zero UI lag. Word files are imported and exported asynchronously with zero loss of unsupported XML elements.

### 2. Selection-Aware Collaboration
- Highlights instantly provide precise context (`block_id`, character offsets, surrounding paragraphs, active section headers, document profile) to the AI Agent.

### 3. Reversible Operation Engine & Multi-Op Transactions
- Every mutation executes through deterministic, serializable, and reversible commands (`InsertTextOp`, `ReplaceTextOp`, `FormatParagraphOp`, `InsertCitationOp`).
- Agent proposals are bundled into single **Agent Transactions** with rich diff previews and one-click full rollback.

### 4. Zero-Hallucination Research Assistant & Citations
- Evaluates empirical claims and attaches verified peer-reviewed sources.
- First-class support for **APA 7th**, **IEEE**, and **Academic-VN** citation and reference list formatting without fabricating DOIs, authors, or publication years.

### 5. Native Diagram & Media Synthesizer
- Generates structured **Mermaid** and clean vector **SVG** diagrams (system architectures, flowcharts, sequence diagrams, and use-case maps) directly into documents.

### 6. Ambiguity & Clarification Engine
- Internal confidence assessment (`HIGH`, `MEDIUM`, `LOW`). If user intent is materially ambiguous (e.g. "rewrite this"), the agent prompts with structured multiple-choice options rather than guessing.

### 7. Dual Verification (Structural + Visual Layout)
- Reopens real `.docx` packages to verify XML validity, typography compliance, and visual layout constraints (overflows, clipped tables, orphan headings, excessive whitespace).

---

## 📦 Installation

```bash
git clone https://github.com/TranNhatThach/Docx-Agent.git
cd Docx-Agent
pip install -e .
```

---

## 🖥️ Usage

### Launch Interactive Visual Workspace
```bash
docx-agent workspace report.docx
```

### Inspect & Read Documents
```bash
# Structural summary
docx-agent inspect report.docx --json

# Read paragraphs with deterministic IDs
docx-agent read report.docx --start 0 --end 10 --json
```

### Format-Preserving Replacement
```bash
docx-agent replace report.docx --target "thuật toán cũ" --replace "mô hình học sâu"
```

### Apply Academic Presets
```bash
docx-agent preset report.docx --name "academic-vn"
```

### Research & Citations
```bash
docx-agent research "Attention Is All You Need Transformer" --style "apa" --json
```

### Generate System Architecture Diagrams
```bash
docx-agent diagram --type architecture --title "Hệ Thống Phân Tán" --item "Client UI" --item "API Gateway" --item "Agent Engine" --item "Vector Database" --json
```

### Dual Verification
```bash
docx-agent visual-verify report.docx --json
docx-agent verify report.docx --expected-font "Times New Roman" --json
```

---

## 🤖 MCP Server Setup (Antigravity / Claude Code / Cursor / Cline)

Add to your agent configuration:

```json
{
  "mcpServers": {
    "docx-agent": {
      "command": "docx-agent-mcp"
    }
  }
}
```

### Available MCP Tools:
- `docx_inspect`: Summary and geometry inspection.
- `docx_read`: Paragraph retrieval with deterministic IDs.
- `docx_selection_context`: Selection-aware surrounding context extractor.
- `docx_research_claim`: Un-hallucinated peer-reviewed citation discovery.
- `docx_generate_diagram`: Mermaid & SVG diagram generation.
- `docx_visual_verify`: Visual layout anomaly detection.
- `docx_clarify`: Ambiguity assessment and multiple-choice options.
- `docx_replace`, `docx_format_text`, `docx_format_paragraph`, `docx_preset`, `docx_table`, `docx_image`, `docx_diff`.

---

## 🧪 Testing & Validation

Run the complete 36-test suite:
```bash
pytest tests/ -v
```

---

## 📄 License

Licensed under the [Apache License 2.0](LICENSE).
