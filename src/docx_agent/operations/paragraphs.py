"""
Paragraph-level formatting operations.
"""

from typing import Optional, Union, Dict, Any, List
import docx
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.paragraph import Paragraph
from docx_agent.core.document import DocumentModel
from docx_agent.core.resolver import TargetResolver


ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "distribute": WD_ALIGN_PARAGRAPH.DISTRIBUTE,
}


class ParagraphFormattingOperations:
    """
    Applies paragraph layout, spacing, indentation, and alignment.
    """

    def __init__(self, model: DocumentModel):
        self.model = model
        self.resolver = TargetResolver(model)

    def format_paragraph(
        self,
        target: Union[str, int, Dict[str, Any]],
        alignment: Optional[str] = None,
        line_spacing: Optional[float] = None,
        space_before_pt: Optional[float] = None,
        space_after_pt: Optional[float] = None,
        first_line_indent_cm: Optional[float] = None,
        left_indent_cm: Optional[float] = None,
        right_indent_cm: Optional[float] = None,
        keep_with_next: Optional[bool] = None,
        page_break_before: Optional[bool] = None,
        widow_control: Optional[bool] = None,
    ) -> int:
        """
        Applies layout and spacing properties to matched paragraph(s).
        Returns count of paragraphs formatted.
        """
        paragraphs = self.resolver.resolve_paragraphs(target)
        for p in paragraphs:
            fmt = p.paragraph_format

            if alignment is not None:
                align_enum = ALIGN_MAP.get(alignment.lower())
                if align_enum is not None:
                    fmt.alignment = align_enum

            if line_spacing is not None:
                fmt.line_spacing = line_spacing

            if space_before_pt is not None:
                fmt.space_before = Pt(space_before_pt)

            if space_after_pt is not None:
                fmt.space_after = Pt(space_after_pt)

            if first_line_indent_cm is not None:
                fmt.first_line_indent = Cm(first_line_indent_cm)

            if left_indent_cm is not None:
                fmt.left_indent = Cm(left_indent_cm)

            if right_indent_cm is not None:
                fmt.right_indent = Cm(right_indent_cm)

            if keep_with_next is not None:
                fmt.keep_with_next = keep_with_next

            if page_break_before is not None:
                fmt.page_break_before = page_break_before

            if widow_control is not None:
                fmt.widow_control = widow_control

        return len(paragraphs)
