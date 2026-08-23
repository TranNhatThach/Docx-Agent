"""
High-Fidelity DOCX Import and Export Adapters.
Bi-directionally converts between Microsoft Word OpenXML packages and Canonical Document Model.
"""

from pathlib import Path
from typing import Union, Optional, List, Dict, Any
import docx
from docx.shared import Cm, Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from docx_agent.canonical.model import (
    DocumentNode,
    SectionNode,
    SectionProperties,
    ParagraphBlock,
    HeadingBlock,
    ListItemBlock,
    TableBlock,
    TableCellNode,
    ImageBlock,
    DiagramBlock,
    UnsupportedBlock,
    RunNode,
    BlockNode,
    generate_id,
)
from docx_agent.ooxml.helpers import set_run_text_preserving_spaces
from docx_agent.ooxml.fields import add_page_number_to_paragraph, add_toc_field
from docx_agent.ooxml.tables import (
    set_cell_background,
    set_row_repeat_header,
    set_row_cant_split,
    set_table_borders,
)
from docx_agent.operations.formatting import hex_to_rgb, HIGHLIGHT_MAP
from docx_agent.operations.paragraphs import ALIGN_MAP
from docx_agent.utils.paths import resolve_safe_path, ensure_parent_dir
from docx_agent.utils.unicode import normalize_unicode


class DocxImporter:
    """
    Parses OpenXML Word documents into the rich Canonical Document Model.
    """

    @staticmethod
    def import_docx(file_path: Union[str, Path]) -> DocumentNode:
        path = resolve_safe_path(file_path)
        doc = docx.Document(str(path))
        doc_node = DocumentNode(title=path.stem)

        doc_node.sections = []

        # Iterate over sections
        for s_idx, sec in enumerate(doc.sections):
            orient = "portrait"
            if sec.page_width > sec.page_height:
                orient = "landscape"

            sec_props = SectionProperties(
                page_width_cm=round(sec.page_width.cm, 2),
                page_height_cm=round(sec.page_height.cm, 2),
                orientation=orient,
                margin_top_cm=round(sec.top_margin.cm, 2),
                margin_bottom_cm=round(sec.bottom_margin.cm, 2),
                margin_left_cm=round(sec.left_margin.cm, 2),
                margin_right_cm=round(sec.right_margin.cm, 2),
            )
            sec_node = SectionNode(id=f"sec_{s_idx+1:04d}", properties=sec_props)

            # Ingest body elements (Paragraphs & Tables)
            # In docx-python, iterate through doc._body elements
            for child in doc._body._element:
                tag = child.tag
                if tag.endswith("p"):
                    # Paragraph element
                    p = docx.text.paragraph.Paragraph(child, doc)
                    style_name = p.style.name if p.style else "Normal"
                    
                    runs_nodes = []
                    for r in p.runs:
                        font_size = r.font.size.pt if r.font and r.font.size else None
                        font_name = r.font.name if r.font else None
                        color_hex = str(r.font.color.rgb) if r.font and r.font.color and r.font.color.rgb else None
                        runs_nodes.append(
                            RunNode(
                                text=normalize_unicode(r.text),
                                font_name=font_name,
                                font_size_pt=font_size,
                                bold=r.bold,
                                italic=r.italic,
                                underline=bool(r.underline),
                                color_rgb=color_hex,
                            )
                        )

                    align_str = str(p.alignment).split(".")[-1].lower() if p.alignment is not None else "justify"
                    ls = p.paragraph_format.line_spacing or 1.5
                    sb = p.paragraph_format.space_before.pt if p.paragraph_format.space_before else 0.0
                    sa = p.paragraph_format.space_after.pt if p.paragraph_format.space_after else 6.0
                    fl = p.paragraph_format.first_line_indent.cm if p.paragraph_format.first_line_indent else 0.0

                    if style_name.startswith("Heading"):
                        try:
                            lvl = int("".join(filter(str.isdigit, style_name)) or 1)
                        except ValueError:
                            lvl = 1
                        sec_node.blocks.append(
                            HeadingBlock(
                                style_name=style_name,
                                level=lvl,
                                runs=runs_nodes,
                                alignment=align_str,
                                line_spacing=ls,
                                space_before_pt=sb,
                                space_after_pt=sa,
                                first_line_indent_cm=0.0,
                            )
                        )
                    elif style_name.startswith("List"):
                        sec_node.blocks.append(
                            ListItemBlock(
                                style_name=style_name,
                                list_type="number" if "Number" in style_name else "bullet",
                                runs=runs_nodes,
                                alignment=align_str,
                                line_spacing=ls,
                            )
                        )
                    else:
                        sec_node.blocks.append(
                            ParagraphBlock(
                                style_name=style_name,
                                runs=runs_nodes,
                                alignment=align_str,
                                line_spacing=ls,
                                space_before_pt=sb,
                                space_after_pt=sa,
                                first_line_indent_cm=fl,
                            )
                        )

                elif tag.endswith("tbl"):
                    # Table element
                    tbl = docx.table.Table(child, doc)
                    num_rows = len(tbl.rows)
                    num_cols = len(tbl.columns) if num_rows else 0
                    cells_grid = []

                    for row in tbl.rows:
                        row_cells = []
                        for cell in row.cells:
                            row_cells.append(
                                TableCellNode(
                                    text=normalize_unicode(cell.text),
                                    runs=[RunNode(text=normalize_unicode(cell.text))],
                                )
                            )
                        cells_grid.append(row_cells)

                    sec_node.blocks.append(
                        TableBlock(
                            rows=num_rows,
                            columns=num_cols,
                            cells=cells_grid,
                            style_name=tbl.style.name if tbl.style else "Table Grid",
                        )
                    )

                elif not (tag.endswith("sectPr")):
                    # Preserve unsupported elements
                    sec_node.blocks.append(
                        UnsupportedBlock(
                            tag_name=tag,
                            raw_xml=child.xml if hasattr(child, "xml") else "",
                        )
                    )

            doc_node.sections.append(sec_node)

        if not doc_node.sections:
            doc_node.sections.append(SectionNode())

        return doc_node


class DocxExporter:
    """
    Serializes Canonical Document Model into an OpenXML Word .docx document.
    """

    @staticmethod
    def export_docx(doc_node: DocumentNode, output_path: Union[str, Path]) -> Path:
        out_p = ensure_parent_dir(output_path)
        doc = docx.Document()

        # Clear default first paragraph
        if doc.paragraphs:
            p_elem = doc.paragraphs[0]._p
            p_elem.getparent().remove(p_elem)

        for s_idx, sec_node in enumerate(doc_node.sections):
            sec = doc.sections[s_idx] if s_idx < len(doc.sections) else doc.add_section()
            
            # Geometry
            sec.page_width = Cm(sec_node.properties.page_width_cm)
            sec.page_height = Cm(sec_node.properties.page_height_cm)
            sec.top_margin = Cm(sec_node.properties.margin_top_cm)
            sec.bottom_margin = Cm(sec_node.properties.margin_bottom_cm)
            sec.left_margin = Cm(sec_node.properties.margin_left_cm)
            sec.right_margin = Cm(sec_node.properties.margin_right_cm)

            # Footer / Page numbering
            if sec_node.has_page_numbers:
                footer = sec.footer
                footer.is_linked_to_previous = False
                p_foot = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
                p_foot.text = ""
                add_page_number_to_paragraph(p_foot, format_str=sec_node.page_number_format)
                p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Header
            if sec_node.header_text:
                header = sec.header
                header.is_linked_to_previous = False
                p_head = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
                p_head.text = sec_node.header_text
                p_head.alignment = WD_ALIGN_PARAGRAPH.RIGHT

            # Blocks
            for blk in sec_node.blocks:
                if isinstance(blk, HeadingBlock):
                    p = doc.add_heading(level=blk.level)
                    p.text = ""
                    for r in blk.runs:
                        run = p.add_run(r.text)
                        if r.bold is not None:
                            run.bold = r.bold
                        if r.italic is not None:
                            run.italic = r.italic
                        if r.font_name:
                            run.font.name = r.font_name
                        if r.font_size_pt:
                            run.font.size = Pt(r.font_size_pt)
                        if r.color_rgb:
                            run.font.color.rgb = hex_to_rgb(r.color_rgb)

                elif isinstance(blk, ListItemBlock):
                    style = "List Bullet" if blk.list_type == "bullet" else "List Number"
                    p = doc.add_paragraph(style=style)
                    for r in blk.runs:
                        run = p.add_run(r.text)
                        if r.bold is not None:
                            run.bold = r.bold
                        if r.italic is not None:
                            run.italic = r.italic

                elif isinstance(blk, ParagraphBlock):
                    p = doc.add_paragraph(style=blk.style_name or "Normal")
                    if blk.alignment:
                        align_enum = ALIGN_MAP.get(blk.alignment.lower())
                        if align_enum is not None:
                            p.alignment = align_enum
                    if blk.line_spacing:
                        p.paragraph_format.line_spacing = blk.line_spacing
                    if blk.space_before_pt is not None:
                        p.paragraph_format.space_before = Pt(blk.space_before_pt)
                    if blk.space_after_pt is not None:
                        p.paragraph_format.space_after = Pt(blk.space_after_pt)
                    if blk.first_line_indent_cm is not None:
                        p.paragraph_format.first_line_indent = Cm(blk.first_line_indent_cm)

                    for r in blk.runs:
                        run = p.add_run(r.text)
                        if r.bold is not None:
                            run.bold = r.bold
                        if r.italic is not None:
                            run.italic = r.italic
                        if r.underline is not None:
                            run.underline = r.underline
                        if r.font_name:
                            run.font.name = r.font_name
                        if r.font_size_pt:
                            run.font.size = Pt(r.font_size_pt)
                        if r.color_rgb:
                            run.font.color.rgb = hex_to_rgb(r.color_rgb)

                elif isinstance(blk, TableBlock):
                    tbl = doc.add_table(rows=blk.rows, cols=blk.columns)
                    try:
                        tbl.style = blk.style_name or "Table Grid"
                    except Exception:
                        tbl.style = "Table Grid"

                    for r_idx, row in enumerate(blk.cells):
                        if r_idx < len(tbl.rows):
                            for c_idx, cell_node in enumerate(row):
                                if c_idx < len(tbl.rows[r_idx].cells):
                                    c = tbl.rows[r_idx].cells[c_idx]
                                    c.text = cell_node.text
                                    if cell_node.bg_color_hex:
                                        set_cell_background(c, cell_node.bg_color_hex)

                    if blk.repeat_header and tbl.rows:
                        set_row_repeat_header(tbl.rows[0], True)
                    for row in tbl.rows:
                        set_row_cant_split(row, True)

                elif isinstance(blk, ImageBlock):
                    if blk.source_uri_or_path and Path(blk.source_uri_or_path).exists():
                        p = doc.add_paragraph()
                        p.alignment = ALIGN_MAP.get(blk.alignment.lower(), WD_ALIGN_PARAGRAPH.CENTER)
                        kwargs = {}
                        if blk.width_cm:
                            kwargs["width"] = Cm(blk.width_cm)
                        p.add_run().add_picture(blk.source_uri_or_path, **kwargs)
                        if blk.caption:
                            cap = doc.add_paragraph(blk.caption, style="Caption")
                            cap.alignment = p.alignment

                elif isinstance(blk, DiagramBlock):
                    # Diagram block output
                    p = doc.add_paragraph(f"[Sơ đồ: {blk.caption or blk.diagram_type.upper()}]")
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if blk.caption:
                        cap = doc.add_paragraph(blk.caption, style="Caption")
                        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER

                elif isinstance(blk, UnsupportedBlock):
                    # Inject raw XML back if preserved
                    if blk.raw_xml:
                        try:
                            elem = parse_xml(blk.raw_xml)
                            doc._body._element.append(elem)
                        except Exception:
                            pass

        doc.save(str(out_p))
        return out_p
