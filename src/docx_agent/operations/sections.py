"""
Page and section layout operations.
"""

from typing import Optional, Union, Dict, Any, List
import docx
from docx.shared import Cm, Inches, Pt
from docx.enum.section import WD_SECTION_START, WD_ORIENT
from docx_agent.core.document import DocumentModel
from docx_agent.core.resolver import TargetResolver


PAGE_SIZES = {
    "a4": (Cm(21.0), Cm(29.7)),
    "letter": (Inches(8.5), Inches(11.0)),
    "legal": (Inches(8.5), Inches(14.0)),
    "a3": (Cm(29.7), Cm(42.0)),
    "a5": (Cm(14.8), Cm(21.0)),
}


class SectionOperations:
    """
    Manages page setup, margins, orientation, and section breaks.
    """

    def __init__(self, model: DocumentModel):
        self.model = model
        self.resolver = TargetResolver(model)

    def configure_section(
        self,
        target: Optional[Union[str, int]] = None,
        page_size: Optional[str] = "a4",
        orientation: str = "portrait",
        margin_top_cm: Optional[float] = None,
        margin_bottom_cm: Optional[float] = None,
        margin_left_cm: Optional[float] = None,
        margin_right_cm: Optional[float] = None,
    ) -> bool:
        """
        Configures geometry, margins, and orientation of a section.
        If target is None, configures all sections in document.
        """
        sections = [self.resolver.resolve_section(target)] if target is not None else list(self.model.doc.sections)

        for sec in sections:
            if page_size and page_size.lower() in PAGE_SIZES:
                w, h = PAGE_SIZES[page_size.lower()]
                if orientation.lower() == "landscape":
                    sec.page_width = h
                    sec.page_height = w
                    sec.orientation = WD_ORIENT.LANDSCAPE
                else:
                    sec.page_width = w
                    sec.page_height = h
                    sec.orientation = WD_ORIENT.PORTRAIT

            if margin_top_cm is not None:
                sec.top_margin = Cm(margin_top_cm)
            if margin_bottom_cm is not None:
                sec.bottom_margin = Cm(margin_bottom_cm)
            if margin_left_cm is not None:
                sec.left_margin = Cm(margin_left_cm)
            if margin_right_cm is not None:
                sec.right_margin = Cm(margin_right_cm)

        return True

    def add_section_break(self, break_type: str = "next_page") -> str:
        """Adds a section break at the end of the document."""
        st_enum = WD_SECTION_START.NEW_PAGE
        if break_type == "continuous":
            st_enum = WD_SECTION_START.CONTINUOUS
        elif break_type == "even_page":
            st_enum = WD_SECTION_START.EVEN_PAGE
        elif break_type == "odd_page":
            st_enum = WD_SECTION_START.ODD_PAGE

        new_sec = self.model.doc.add_section(st_enum)
        self.model.reindex()
        return f"sec_{len(self.model.doc.sections):04d}"
