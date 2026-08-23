"""
OOXML table helpers for cell shading, borders, and header repetition.
"""

from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.table import Table, _Row, _Cell


def set_cell_background(cell: _Cell, hex_color: str) -> None:
    """Sets cell background shading color (e.g. 'F0F0F0')."""
    clean_hex = hex_color.lstrip("#")
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{clean_hex}"/>'
    tcPr = cell._tc.get_or_add_tcPr()
    # Remove existing shd if any
    for existing in tcPr.xpath("./w:shd"):
        tcPr.remove(existing)
    tcPr.append(parse_xml(shading_xml))


def set_row_repeat_header(row: _Row, val: bool = True) -> None:
    """Marks row as repeating header across page breaks (w:tblHeader)."""
    trPr = row._tr.get_or_add_trPr()
    for existing in trPr.xpath("./w:tblHeader"):
        trPr.remove(existing)
    if val:
        trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))


def set_row_cant_split(row: _Row, val: bool = True) -> None:
    """Prevents row from splitting across page breaks (w:cantSplit)."""
    trPr = row._tr.get_or_add_trPr()
    for existing in trPr.xpath("./w:cantSplit"):
        trPr.remove(existing)
    if val:
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))


def set_table_borders(table: Table, color: str = "D3D3D3", sz: str = "4", val: str = "single") -> None:
    """Sets subtle clean borders on the table."""
    clean_color = color.lstrip("#")
    tblPr = table._tbl.tblPr
    for existing in tblPr.xpath("./w:tblBorders"):
        tblPr.remove(existing)
    borders_xml = f"""
    <w:tblBorders {nsdecls("w")}>
        <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{clean_color}"/>
        <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{clean_color}"/>
        <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{clean_color}"/>
        <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{clean_color}"/>
        <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{clean_color}"/>
        <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{clean_color}"/>
    </w:tblBorders>
    """
    tblPr.append(parse_xml(borders_xml))
