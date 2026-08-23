"""
Deterministic identity assignment and registry for document elements.
"""

from typing import Dict, Any, Optional, List, Tuple
import docx
from docx.text.paragraph import Paragraph
from docx.table import Table, _Cell, _Row
from docx.section import Section
from docx_agent.core.elements import ElementType


class IdentityManager:
    """
    Manages deterministic IDs and index mappings for all elements within a document.
    """

    def __init__(self, doc: docx.Document):
        self.doc = doc
        self.p_map: Dict[str, Paragraph] = {}
        self.p_id_by_obj: Dict[int, str] = {}
        self.p_idx_by_id: Dict[str, int] = {}
        
        self.tbl_map: Dict[str, Table] = {}
        self.tbl_id_by_obj: Dict[int, str] = {}
        self.tbl_idx_by_id: Dict[str, int] = {}

        self.cell_map: Dict[str, _Cell] = {}
        self.sec_map: Dict[str, Section] = {}

        self.reindex()

    def reindex(self) -> None:
        """Re-indexes all elements in the document."""
        self.p_map.clear()
        self.p_id_by_obj.clear()
        self.p_idx_by_id.clear()
        self.tbl_map.clear()
        self.tbl_id_by_obj.clear()
        self.tbl_idx_by_id.clear()
        self.cell_map.clear()
        self.sec_map.clear()

        # Paragraphs
        for idx, p in enumerate(self.doc.paragraphs):
            pid = f"p_{idx + 1:04d}"
            self.p_map[pid] = p
            self.p_id_by_obj[id(p._p)] = pid
            self.p_idx_by_id[pid] = idx

        # Tables
        for t_idx, table in enumerate(self.doc.tables):
            tid = f"tbl_{t_idx + 1:04d}"
            self.tbl_map[tid] = table
            self.tbl_id_by_obj[id(table._tbl)] = tid
            self.tbl_idx_by_id[tid] = t_idx

            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    cid = f"{tid}_r{r_idx + 1:02d}_c{c_idx + 1:02d}"
                    self.cell_map[cid] = cell

        # Sections
        for s_idx, sec in enumerate(self.doc.sections):
            sid = f"sec_{s_idx + 1:04d}"
            self.sec_map[sid] = sec

    def get_paragraph(self, target_id: str) -> Optional[Paragraph]:
        return self.p_map.get(target_id)

    def get_table(self, target_id: str) -> Optional[Table]:
        return self.tbl_map.get(target_id)

    def get_cell(self, target_id: str) -> Optional[_Cell]:
        return self.cell_map.get(target_id)

    def get_section(self, target_id: str) -> Optional[Section]:
        return self.sec_map.get(target_id)

    def get_paragraph_id(self, p: Paragraph) -> Optional[str]:
        return self.p_id_by_obj.get(id(p._p))

    def get_table_id(self, table: Table) -> Optional[str]:
        return self.tbl_id_by_obj.get(id(table._tbl))
