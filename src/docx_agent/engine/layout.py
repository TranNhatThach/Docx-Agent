"""
LayoutEngine: Deterministic WordprocessingML Geometry, Page Breaking & Pagination Engine.
Converts Canonical Document Model into a Paginated Render Tree (A4 Pages -> Headers -> Content -> Footers).
"""

import math
from typing import List, Dict, Any, Optional
from pydantic import BaseModel, Field

from docx_agent.canonical.model import (
    DocumentNode,
    SectionNode,
    ParagraphBlock,
    HeadingBlock,
    ListItemBlock,
    TableBlock,
    ImageBlock,
    DiagramBlock,
    UnsupportedBlock,
    BlockNode,
)


class PageGeometry(BaseModel):
    width_pt: float = 595.28  # A4 210mm
    height_pt: float = 841.89  # A4 297mm
    margin_top_pt: float = 56.69  # 2.0cm
    margin_bottom_pt: float = 56.69  # 2.0cm
    margin_left_pt: float = 85.04  # 3.0cm
    margin_right_pt: float = 56.69  # 2.0cm
    content_width_pt: float = 453.55  # 16.0cm
    content_height_pt: float = 728.51  # 25.7cm


class LayoutPage(BaseModel):
    page_number: int
    total_pages: int = 1
    section_index: int = 0
    is_first_page_of_section: bool = False
    is_cover_page: bool = False
    has_header: bool = True
    has_footer: bool = True
    header_text: Optional[str] = None
    footer_left_text: Optional[str] = None
    footer_right_text: Optional[str] = None
    geometry: PageGeometry = Field(default_factory=PageGeometry)
    blocks: List[Dict[str, Any]] = Field(default_factory=list)


class LayoutDocument(BaseModel):
    title: str = "Document"
    total_pages: int = 1
    total_words: int = 0
    total_chars: int = 0
    pages: List[LayoutPage] = Field(default_factory=list)
    headings_outline: List[Dict[str, Any]] = Field(default_factory=list)


class LayoutEngine:
    """
    High-Fidelity Deterministic Pagination and Layout Engine.
    """

    CM_TO_PT = 28.3465
    AVG_CHAR_WIDTH_RATIO = 0.52

    @classmethod
    def paginate(cls, doc_node: DocumentNode) -> LayoutDocument:
        pages: List[LayoutPage] = []
        headings_outline: List[Dict[str, Any]] = []
        total_words = 0
        total_chars = 0
        current_page_num = 1

        for s_idx, sec in enumerate(doc_node.sections):
            geom = PageGeometry(
                width_pt=sec.properties.page_width_cm * cls.CM_TO_PT,
                height_pt=sec.properties.page_height_cm * cls.CM_TO_PT,
                margin_top_pt=sec.properties.margin_top_cm * cls.CM_TO_PT,
                margin_bottom_pt=sec.properties.margin_bottom_cm * cls.CM_TO_PT,
                margin_left_pt=sec.properties.margin_left_cm * cls.CM_TO_PT,
                margin_right_pt=sec.properties.margin_right_cm * cls.CM_TO_PT,
                content_width_pt=(sec.properties.page_width_cm - sec.properties.margin_left_cm - sec.properties.margin_right_cm) * cls.CM_TO_PT,
                content_height_pt=(sec.properties.page_height_cm - sec.properties.margin_top_cm - sec.properties.margin_bottom_cm) * cls.CM_TO_PT,
            )

            current_page = LayoutPage(
                page_number=current_page_num,
                section_index=s_idx,
                is_first_page_of_section=True,
                is_cover_page=(current_page_num == 1 and sec.properties.different_first_page),
                geometry=geom,
            )
            current_page_height = 0.0

            for b_idx, blk in enumerate(sec.blocks):
                # Count statistics
                if hasattr(blk, "full_text"):
                    txt = blk.full_text
                    total_words += len(txt.split())
                    total_chars += len(txt)
                elif isinstance(blk, TableBlock):
                    for row in blk.cells:
                        for cell in row:
                            total_words += len(cell.text.split())
                            total_chars += len(cell.text)

                # Record Outline
                if isinstance(blk, HeadingBlock):
                    headings_outline.append({
                        "id": blk.id,
                        "level": blk.level,
                        "text": blk.full_text,
                        "style": blk.style_name or f"Heading {blk.level}",
                    })
                elif isinstance(blk, ParagraphBlock) and blk.style_name:
                    s_low = blk.style_name.lower().strip()
                    if s_low in ("title", "tựa đề"):
                        headings_outline.append({
                            "id": blk.id,
                            "level": 0,
                            "text": blk.full_text,
                            "style": blk.style_name,
                        })
                    elif s_low in ("subtitle", "phụ đề"):
                        headings_outline.append({
                            "id": blk.id,
                            "level": 1,
                            "text": blk.full_text,
                            "style": blk.style_name,
                        })

                # Check explicit page break triggers
                has_page_break = False
                if isinstance(blk, ParagraphBlock):
                    if blk.page_break_before:
                        has_page_break = True
                    for r in blk.runs:
                        if r.is_page_break or "\x0c" in r.text:
                            has_page_break = True

                # Estimate block height
                blk_height = cls._estimate_block_height(blk, geom.content_width_pt)

                # Check if block fits on current page
                should_break_page = False
                if has_page_break and len(current_page.blocks) > 0:
                    should_break_page = True
                elif current_page_height + blk_height > geom.content_height_pt and len(current_page.blocks) > 0:
                    should_break_page = True
                elif getattr(blk, "keep_with_next", False) and b_idx + 1 < len(sec.blocks):
                    next_blk = sec.blocks[b_idx + 1]
                    next_h = cls._estimate_block_height(next_blk, geom.content_width_pt)
                    if current_page_height + blk_height + min(next_h, 40.0) > geom.content_height_pt and len(current_page.blocks) > 0:
                        should_break_page = True

                if should_break_page:
                    pages.append(current_page)
                    current_page_num += 1
                    current_page = LayoutPage(
                        page_number=current_page_num,
                        section_index=s_idx,
                        is_first_page_of_section=False,
                        is_cover_page=False,
                        geometry=geom,
                    )
                    current_page_height = 0.0

                # Serialize block dictionary with resolved layout attributes
                blk_dict = cls._serialize_block_for_render(blk)
                current_page.blocks.append(blk_dict)
                current_page_height += blk_height

            if current_page.blocks or len(pages) == 0:
                pages.append(current_page)
                current_page_num += 1

        total_pages = max(1, len(pages))
        for p in pages:
            p.total_pages = total_pages
            sec_props = doc_node.sections[p.section_index].properties
            sec_node = doc_node.sections[p.section_index]

            if p.page_number == 1 and sec_props.different_first_page:
                p.is_cover_page = True
                p.has_header = bool(sec_node.first_page_header_text)
                p.has_footer = bool(sec_node.first_page_footer_text)
                p.header_text = sec_node.first_page_header_text
                p.footer_left_text = sec_node.first_page_footer_text
                p.footer_right_text = None
            else:
                p.is_cover_page = False
                p.has_header = bool(sec_node.header_text)
                p.has_footer = sec_node.has_page_numbers or bool(sec_node.footer_text)
                p.header_text = sec_node.header_text
                p.footer_left_text = sec_node.footer_text or "Trường Đại học Giao thông Vận tải (UTC)"
                p.footer_right_text = f"Trang {p.page_number} / {total_pages}"

        return LayoutDocument(
            title=doc_node.title,
            total_pages=total_pages,
            total_words=total_words,
            total_chars=total_chars,
            pages=pages,
            headings_outline=headings_outline,
        )

    @classmethod
    def _estimate_block_height(cls, blk: BlockNode, content_width_pt: float) -> float:
        if isinstance(blk, TableBlock):
            total_tbl_h = 0.0
            for row in blk.cells:
                max_row_h = 20.0
                num_cols = max(1, len(row))
                col_w_pt = content_width_pt / num_cols
                for cell in row:
                    txt = cell.text or ""
                    lines = txt.split("\n")
                    if "SELECT" in txt or "FROM" in txt or "WHERE" in txt:
                        font_size = 10.0
                        line_h = 12.0
                        char_w = 6.0
                        chars_per_line = max(10, int((col_w_pt - 16.0) / char_w))
                        total_lines = sum(max(1, math.ceil(len(l) / chars_per_line)) for l in lines)
                        cell_h = (total_lines * line_h) + 16.0 + 4.0
                    elif "Nhận xét" in txt or "Đánh giá" in txt:
                        font_size = 12.0
                        line_h = 15.6
                        char_w = 5.5
                        chars_per_line = max(10, int((col_w_pt - 16.0) / char_w))
                        total_lines = sum(max(1, math.ceil(len(l) / chars_per_line)) for l in lines)
                        cell_h = (total_lines * line_h) + 16.0 + 8.0
                    else:
                        font_size = 11.0
                        line_h = 14.0
                        char_w = 5.5
                        chars_per_line = max(10, int((col_w_pt - 12.0) / char_w))
                        total_lines = sum(max(1, math.ceil(len(l) / chars_per_line)) for l in lines)
                        cell_h = (total_lines * line_h) + 12.0
                    if cell_h > max_row_h:
                        max_row_h = cell_h
                total_tbl_h += max_row_h
            return total_tbl_h + 4.0

        elif isinstance(blk, ParagraphBlock):
            font_size = 13.0
            line_spacing = blk.line_spacing or 1.4
            sb = blk.space_before_pt or 0.0
            sa = blk.space_after_pt or 6.0

            if isinstance(blk, HeadingBlock):
                font_size = 16.0 if blk.level == 1 else (14.0 if blk.level == 2 else 13.0)
                line_spacing = 1.2
                sb = 16.0 if blk.level == 1 else (12.0 if blk.level == 2 else 8.0)
                sa = 8.0 if blk.level == 1 else (6.0 if blk.level == 2 else 4.0)
            elif blk.runs and blk.runs[0].font_size_pt:
                font_size = blk.runs[0].font_size_pt

            line_height_pt = font_size * line_spacing
            char_width_pt = font_size * 0.46
            chars_per_line = max(20, int(content_width_pt / char_width_pt))

            lines = blk.full_text.split("\n")
            total_lines = sum(max(1, math.ceil(len(l) / chars_per_line)) for l in lines)
            return (total_lines * line_height_pt) + sb + sa

        elif isinstance(blk, (ImageBlock, DiagramBlock)):
            width_pt = (blk.width_cm or 12.0) * cls.CM_TO_PT
            aspect = getattr(blk, "aspect_ratio", None) or 0.65
            return (width_pt * aspect) + 30.0

        return 24.0

    @classmethod
    def _serialize_block_for_render(cls, blk: BlockNode) -> Dict[str, Any]:
        btype = blk.type.value if hasattr(blk.type, "value") else str(blk.type)
        res = {
            "id": blk.id,
            "type": btype,
            "style_name": getattr(blk, "style_name", "Normal"),
        }

        if isinstance(blk, (ParagraphBlock, HeadingBlock, ListItemBlock)):
            res.update({
                "text": blk.full_text,
                "alignment": blk.alignment or "justify",
                "line_spacing": blk.line_spacing or 1.4,
                "space_before_pt": blk.space_before_pt or 0.0,
                "space_after_pt": blk.space_after_pt or 6.0,
                "first_line_indent_cm": blk.first_line_indent_cm or 0.0,
                "left_indent_cm": blk.left_indent_cm or 0.0,
                "hanging_indent_cm": getattr(blk, "hanging_indent_cm", 0.0),
                "keep_with_next": getattr(blk, "keep_with_next", False),
                "resolved_numbering_label": getattr(blk, "resolved_numbering_label", None),
                "runs": [r.model_dump() for r in blk.runs],
            })
            if isinstance(blk, HeadingBlock):
                res["level"] = blk.level
            if isinstance(blk, ListItemBlock):
                res["list_type"] = blk.list_type
                res["list_level"] = blk.list_level

        elif isinstance(blk, TableBlock):
            rows_data = []
            for row in blk.cells:
                row_cells = []
                for cell in row:
                    row_cells.append({
                        "id": cell.id,
                        "text": cell.text,
                        "bg_color_hex": cell.bg_color_hex,
                        "rowspan": cell.rowspan,
                        "colspan": cell.colspan,
                        "borders": cell.borders,
                        "vertical_align": cell.vertical_align,
                    })
                rows_data.append(row_cells)

            res.update({
                "rows": blk.rows,
                "columns": blk.columns,
                "cells": rows_data,
                "style_name": blk.style_name,
                "alignment": blk.alignment,
                "repeat_header": blk.repeat_header,
                "col_widths_cm": blk.col_widths_cm,
                "grid_cols_cm": blk.grid_cols_cm,
            })

        elif isinstance(blk, (ImageBlock, DiagramBlock)):
            res.update({
                "source_uri_or_path": getattr(blk, "source_uri_or_path", ""),
                "rendered_svg": getattr(blk, "rendered_svg", None),
                "width_cm": getattr(blk, "width_cm", 10.0),
                "height_cm": getattr(blk, "height_cm", None),
                "alignment": blk.alignment,
                "caption": blk.caption,
            })

        elif isinstance(blk, UnsupportedBlock):
            res.update({
                "tag_name": blk.tag_name,
                "warning_message": blk.warning_message,
                "raw_xml": blk.raw_xml,
            })

        return res
