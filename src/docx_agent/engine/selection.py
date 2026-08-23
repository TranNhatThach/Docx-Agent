"""
Selection Model and Context Extractor for Selection-Aware Agent Interactions.
"""

from typing import Dict, Any, List, Optional
from pydantic import BaseModel, Field

from docx_agent.canonical.model import DocumentNode, SectionNode, BaseBlockNode, ParagraphBlock, HeadingBlock


class SelectionRange(BaseModel):
    start: int = 0
    end: int = 0


class SelectionContext(BaseModel):
    document_id: str
    document_title: str
    document_profile: str
    section_id: str
    section_title: str
    block_id: str
    block_type: str
    block_style: str
    selection: SelectionRange
    selected_text: str
    surrounding_context: Dict[str, Any] = Field(default_factory=dict)
    active_citations: List[Dict[str, Any]] = Field(default_factory=list)


class SelectionProvider:
    """
    Extracts precise context from user highlighting for agent consumption.
    """

    @staticmethod
    def build_context(
        doc: DocumentNode,
        block_id: str,
        start_offset: int,
        end_offset: int,
    ) -> SelectionContext:
        target_blk: Optional[BaseBlockNode] = None
        target_sec: Optional[SectionNode] = None
        blk_idx: int = -1
        current_section_title: str = "Introduction"

        # Locate section and block
        for sec in doc.sections:
            for idx, b in enumerate(sec.blocks):
                if isinstance(b, HeadingBlock):
                    current_section_title = b.full_text
                if b.id == block_id:
                    target_blk = b
                    target_sec = sec
                    blk_idx = idx
                    break
            if target_blk:
                break

        if not target_blk or not target_sec:
            raise ValueError(f"Block ID '{block_id}' not found in document.")

        # Extract selected text
        full_text = getattr(target_blk, "full_text", getattr(target_blk, "text", ""))
        safe_start = max(0, min(start_offset, len(full_text)))
        safe_end = max(safe_start, min(end_offset, len(full_text)))
        selected_substring = full_text[safe_start:safe_end]

        # Surrounding blocks context
        prev_text = ""
        next_text = ""
        if blk_idx > 0:
            prev_b = target_sec.blocks[blk_idx - 1]
            prev_text = getattr(prev_b, "full_text", getattr(prev_b, "text", ""))
        if blk_idx + 1 < len(target_sec.blocks):
            next_b = target_sec.blocks[blk_idx + 1]
            next_text = getattr(next_b, "full_text", getattr(next_b, "text", ""))

        # Active citations
        active_cites = []
        if isinstance(target_blk, ParagraphBlock):
            for r in target_blk.runs:
                if r.citation_id and r.citation_id in doc.citations:
                    c_node = doc.citations[r.citation_id]
                    src = doc.sources.get(c_node.source_id)
                    active_cites.append({
                        "citation_id": c_node.id,
                        "formatted": c_node.formatted_intext,
                        "source_title": src.title if src else "Unknown",
                    })

        return SelectionContext(
            document_id=doc.id,
            document_title=doc.title,
            document_profile=doc.profile.value,
            section_id=target_sec.id,
            section_title=current_section_title,
            block_id=target_blk.id,
            block_type=target_blk.type.value,
            block_style=target_blk.style_name or "Normal",
            selection=SelectionRange(start=safe_start, end=safe_end),
            selected_text=selected_substring,
            surrounding_context={
                "previous_block": prev_text[:200],
                "next_block": next_text[:200],
                "full_block_text": full_text,
            },
            active_citations=active_cites,
        )
