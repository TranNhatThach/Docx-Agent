# Canonical Document Model Architecture

## 1. Overview
In Docx-Agent V2, raw DOCX XML is no longer the runtime editor state. Instead, runtime manipulations occur against the **Canonical Document Model** (`docx_agent.canonical.model`).

## 2. Node Hierarchy

```
DocumentNode (id, title, profile, version, metadata)
 ├── SectionNode (id, properties: page size, orientation, margins, columns)
 │    ├── Header (optional block)
 │    ├── Footer (optional block, page numbers)
 │    └── Blocks (List[BlockNode])
 │         ├── HeadingBlock (level 1-6)
 │         ├── ParagraphBlock (runs, alignment, line spacing, indents)
 │         ├── ListItemBlock (bullet/number, level)
 │         ├── TableBlock (rows, cols, cells, cantSplit, tblHeader)
 │         ├── ImageBlock (source, width, height, alignment, caption)
 │         ├── DiagramBlock (type, source Mermaid, rendered SVG, caption)
 │         ├── CitationBlock (source reference, in-text citation)
 │         └── UnsupportedBlock (tag_name, raw_xml for loss-less preservation)
 └── References & Citations
      ├── Sources (Dict[src_id, SourceMetadata])
      └── Citations (Dict[cite_id, CitationNode])
```

## 3. Benefits
- **Sub-millisecond latency**: In-memory manipulations avoid ZIP repacking overhead on every keystroke.
- **Graceful degradation**: Unsupported Word XML tags are captured as `UnsupportedBlock` and serialized back intact upon DOCX export.
- **Deterministic Identity**: Every node has a stable unique ID (`blk_1234abcd`, `r_5678ef01`, `cite_9012gh34`).
