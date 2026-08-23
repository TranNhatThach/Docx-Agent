# Technical Debt Tracker

This document tracks known technical debt items, architectural constraints, and their prioritized remediation roadmap.

| ID | Component | Problem Description | Impact | Risk | Priority | Recommended Solution |
|---|---|---|---|---|---|---|
| **TD-001** | `engine/layout.py` | Complex multi-column newspaper layouts use single-column bounding approximations | Minor layout discrepancy for multi-column magazines | Low | Low | Implement column-aware flow pagination in `LayoutEngine` |
| **TD-002** | `adapters/docx.py` | Embedded Excel charts (.xlsx OLE parts) are preserved but not dynamically editable via CLI | Charts render as images without chart-data editing | Low | Medium | Add OLE / ChartML data parser in `docx_agent.ooxml.charts` |
| **TD-003** | `extensions/vscode` | Webview bundles CDN scripts for JSZip & docx-preview | Requires internet connection on initial load if not cached | Low | Medium | Bundle minified local scripts into `extensions/vscode/media/vendor/` |
| **TD-004** | `research/citations` | External DOI lookups rely on standard CrossRef public endpoint | Rate-limiting possible during massive batch citation checks | Low | Low | Add local SQLite citation cache and configurable email polite-pool in `core/config.py` |
