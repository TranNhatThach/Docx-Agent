# Docx-Agent System Overview & Architecture Map

## 1. High-Level Architectural Flow
```text
┌────────────────────────────────────────────────────────────────────────┐
│                        INTERFACES LAYER                                │
│  • CLI (Typer)               • MCP Server (FastMCP)                    │
│  • Visual Workspace Server   • VS Code Extension Webview (app.html)    │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │
                                    ▼
┌────────────────────────────────────────────────────────────────────────┐
│                        APPLICATION LAYER                               │
│  • DocumentAgent (Orchestration & High-level APIs)                     │
│  • TransactionContext (Atomic Snapshots, Rollback, Undo/Redo)          │
│  • WorkspaceBridge (IPC & Layout Serialization)                        │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │
                                    ▼
┌────────────────────────────────────────────────────────────────────────┐
│                        DOMAIN / CORE LAYER                             │
│  • Canonical Document Model (DocumentNode, SectionNode, Blocks, Runs)  │
│  • Configuration (Settings, Environment overrides)                     │
│  • Error Taxonomy & Machine-Readable ErrorCode                         │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │
                                    ▼
┌────────────────────────────────────────────────────────────────────────┐
│                        ENGINE & OPERATIONS                             │
│  • LayoutEngine (Deterministic A4 Pagination & Heading Outline)        │
│  • StyleResolver (6-Level Cascading Inheritance)                       │
│  • NumberingResolver (Multilevel Abstract List Evaluation)             │
│  • DocumentValidator & DiffEngine                                      │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │
                                    ▼
┌────────────────────────────────────────────────────────────────────────┐
│                    INFRASTRUCTURE & ADAPTERS                           │
│  • DocxImporter & DocxExporter (python-docx + raw oxml DOM)            │
│  • MarkdownToDocxConverter                                             │
│  • Media / DrawingML Helpers & Unicode Normalizers                     │
└────────────────────────────────────────────────────────────────────────┘
```

## 2. Component Directory Map
- `src/docx_agent/canonical/`: Immutable and mutable nodes representing paragraphs, headings, tables, drawings.
- `src/docx_agent/core/`: Exceptions, error codes, application settings, constants.
- `src/docx_agent/engine/`: Pagination, cascading style resolution, numbering, diff calculations.
- `src/docx_agent/adapters/`: Input/Output serializers for `.docx` OpenXML packages.
- `src/docx_agent/interfaces/`: User/agent entry points (CLI, MCP, HTTP workspace).
- `src/docx_agent/presets/`: Built-in standardized document formatting profiles (Academic VN, IEEE, Corporate).
