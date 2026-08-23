# V2 UPGRADE ARCHITECTURE AUDIT: AI-NATIVE DOCUMENT WORKSPACE

**Date:** 2026-08-23  
**Status:** Phase 0 Baseline Audit Completed  
**Target Version:** v2.0.0  
**Repository:** `https://github.com/TranNhatThach/Docx-Agent.git`

---

## 1. Executive Baseline & Upgrade Goal

Docx-Agent v1.0 establishes an agent-native engine for direct Microsoft Word manipulation with deterministic IDs, run surgery, academic presets, atomic transactions, and structural verification.

In **v2.0**, the project transforms from an **offline file mutation tool** into an **AI-Native Document Workspace + Document Engine**:
- **Dual Representation**: Decoupling runtime document state from raw DOCX XML via a high-performance **Canonical Document Model**. DOCX becomes an interchange/persistence format.
- **Visual Workspace & Human-in-the-Loop**: Interactive document canvas (A4 layout), selection-aware agent chat, live change previews, and multi-operation transaction approval/rejection.
- **Research & Citation Pipeline**: Real source discovery, evidence evaluation, provenance tracking, and bibliography management (APA, IEEE, Academic-VN) without factual fabrication.
- **Intelligent Diagram & Image Engine**: Native diagram synthesis (Mermaid, SVG, architecture, flowcharts) and image provenance.
- **Dual Verification**: Structural schema verification + visual layout verification (overflow detection, typography consistency, spacing sanity).
- **Zero Regressions**: 100% backward compatibility for all existing v1.0 CLI commands, MCP tools, and Python SDK APIs.

---

## 2. Inventory of Reusable V1 Components

| Component | Status in v2 | Rationale & Evolution |
| :--- | :--- | :--- |
| `docx_agent.ooxml.runs` | **Retained & Reused** | Core format-preserving run surgery algorithm will power the DOCX Export Adapter. |
| `docx_agent.ooxml.fields` | **Retained & Reused** | Powers dynamic `PAGE`, `NUMPAGES`, and `TOC` serialization. |
| `docx_agent.ooxml.tables` | **Retained & Reused** | Cell shading, `w:tblHeader`, and `w:cantSplit` generation. |
| `docx_agent.core.identity` | **Extended** | Extended from transient indices to persistent UUIDs across document nodes. |
| `docx_agent.presets` | **Extended** | Presets will apply directly to the Canonical Document Model and style definitions. |
| `docx_agent.transactions` | **Extended** | Elevated to support multi-operation Agent Transactions with undo/redo stacks. |
| `docx_agent.verification` | **Extended** | Augmented with visual layout verification (overflow, blank page detection). |
| `docx_agent.interfaces.cli` | **Preserved & Extended**| Retains all existing commands; adds workspace launch and visual verification. |
| `docx_agent.interfaces.mcp` | **Preserved & Extended**| Retains existing tool schemas; adds selection-aware and research tools. |

---

## 3. Identified Architectural Bottlenecks in V1

1. **Tight Coupling to DOCX XML as Runtime State**:
   - In v1, operations modify `python-docx` XML DOM directly. For an interactive visual editor, serializing and re-reading the entire ZIP archive on each edit causes unacceptable latency.
2. **Lack of In-Memory Command / Reversible Operation Engine**:
   - V1 had atomic transaction rollback via `.bak` file restore. An interactive editor requires fine-grained operational undo/redo (`InsertTextOp`, `FormatOp`, `AgentTransactionOp`) in memory.
3. **No Selection Context DTO**:
   - Agents in v1 operated on global selectors (`p_0012` or text query). V2 requires rich selection context (exact character range, surrounding block context, section title, active citations).
4. **Missing Research & Citation Subsystem**:
   - V1 lacked provenance metadata, citation data structures, and research provider interfaces.
5. **Missing Visual Layout Verification**:
   - V1 checked XML integrity and paragraph attributes but lacked layout-level checks (e.g. image dimension overflow, table clipping, orphan headings).

---

## 4. Target V2 Modular Architecture

```
                          AI CODING AGENTS & HUMAN USER
                                        │
           ┌────────────────────────────┼────────────────────────────┐
           ▼                            ▼                            ▼
     Visual Editor               MCP Interface                 CLI Interface
 (Webview Canvas & Chat)       (docx-agent-mcp)              (docx-agent CLI)
           │                            │                            │
           └────────────────────────────┼────────────────────────────┘
                                        │
                             ┌──────────▼──────────┐
                             │    Document Engine  │
                             │  (Workspace Manager)│
                             └──────────┬──────────┘
                                        │
         ┌──────────────────────────────┼──────────────────────────────┐
         ▼                              ▼                              ▼
┌──────────────────┐          ┌──────────────────┐          ┌──────────────────┐
│ Canonical Model  │          │ Operation Engine │          │ Agent Subsystems │
│ - DocumentNode   │          │ - Command Stack  │          │ - SelectionAgent │
│ - SectionNode    │          │ - Undo/Redo      │          │ - Clarification  │
│ - BlockNode      │          │ - AgentTxManager │          │ - ResearchAgent  │
│ - Inline/RunNode │          │ - Dirty Tracker  │          │ - Citations & Prov│
│ - CitationNode   │          │ - Snapshots      │          │ - Diagram/Image  │
└────────┬─────────┘          └────────┬─────────┘          └────────┬─────────┘
         │                             │                             │
         └─────────────────────────────┼─────────────────────────────┘
                                       │
                    ┌──────────────────┴──────────────────┐
                    ▼                                     ▼
         ┌─────────────────────┐               ┌─────────────────────┐
         │ DOCX Import/Export  │               │ Dual Verification   │
         │ - DocxImporter      │               │ - StructuralCheck   │
         │ - DocxExporter      │               │ - VisualLayoutCheck │
         │ - Run Surgery Bridge│               │ - SemanticDiff      │
         └──────────┬──────────┘               └─────────────────────┘
                    │
                    ▼
          Microsoft Word (.docx)
```

---

## 5. Migration & Phased Rollout Plan

- **Phase 1**: Canonical Document Model (`docx_agent.canonical.model`).
- **Phase 2**: Document Operation & Command Engine (`docx_agent.engine.operations`).
- **Phase 3**: Transaction Manager, Operation Log, Undo/Redo Stack (`docx_agent.engine.transactions`).
- **Phase 4**: Selection Model & Context Provider (`docx_agent.engine.selection`).
- **Phase 5**: Research Assistant & Citation Provenance (`docx_agent.research.citations`).
- **Phase 6**: Diagram & Image Synthesizer (`docx_agent.media.diagrams`).
- **Phase 7**: Clarification & Ambiguity Engine (`docx_agent.agent.clarification`).
- **Phase 8**: DOCX Import & Export Adapters (`docx_agent.adapters.docx`).
- **Phase 9**: Dual Verification (Structural + Visual Layout) (`docx_agent.verification.visual`).
- **Phase 10**: Interactive Visual Workspace (`docx_agent.interfaces.workspace`).
- **Phase 11**: Extended MCP & CLI Tooling (`docx_agent.interfaces.mcp`, `cli`).
- **Phase 12**: Comprehensive Test Suite (Scenarios A through J) & E2E Validation.
