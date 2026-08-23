"""
Style Engine: Inspect, create, update, and apply Word paragraph and character styles.
"""

from typing import Optional, Union, Dict, Any, List
import docx
from docx.shared import Pt, RGBColor, Cm
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx_agent.core.document import DocumentModel
from docx_agent.core.resolver import TargetResolver
from docx_agent.core.exceptions import StyleError
from docx_agent.operations.formatting import hex_to_rgb
from docx_agent.operations.paragraphs import ALIGN_MAP


class StyleOperations:
    """
    Manages document styles and promotes style-driven consistency.
    """

    def __init__(self, model: DocumentModel):
        self.model = model
        self.resolver = TargetResolver(model)

    def list_styles(self) -> List[Dict[str, Any]]:
        """Returns all defined styles in the document."""
        styles_list = []
        for s in self.model.doc.styles:
            # Only list paragraph or character styles
            if s.type in (WD_STYLE_TYPE.PARAGRAPH, WD_STYLE_TYPE.CHARACTER):
                font_name = s.font.name if s.font else None
                font_size = s.font.size.pt if s.font and s.font.size else None
                styles_list.append(
                    {
                        "name": s.name,
                        "type": "paragraph" if s.type == WD_STYLE_TYPE.PARAGRAPH else "character",
                        "font_name": font_name,
                        "font_size_pt": font_size,
                        "builtin": s.builtin,
                    }
                )
        return sorted(styles_list, key=lambda x: x["name"])

    def inspect_style(self, style_name: str) -> Dict[str, Any]:
        """Inspects properties of a specific style."""
        try:
            s = self.model.doc.styles[style_name]
        except KeyError:
            raise StyleError(style_name, "Style not found in document styles gallery.")

        font_info = {}
        if s.font:
            font_info = {
                "name": s.font.name,
                "size_pt": s.font.size.pt if s.font.size else None,
                "bold": s.font.bold,
                "italic": s.font.italic,
                "underline": bool(s.font.underline) if s.font.underline is not None else None,
                "color_rgb": str(s.font.color.rgb) if s.font.color and s.font.color.rgb else None,
            }

        para_info = {}
        if s.type == WD_STYLE_TYPE.PARAGRAPH and hasattr(s, "paragraph_format"):
            fmt = s.paragraph_format
            para_info = {
                "line_spacing": fmt.line_spacing,
                "space_before_pt": fmt.space_before.pt if fmt.space_before else None,
                "space_after_pt": fmt.space_after.pt if fmt.space_after else None,
                "alignment": str(fmt.alignment).split(".")[-1].lower() if fmt.alignment is not None else None,
            }

        return {
            "name": s.name,
            "type": "paragraph" if s.type == WD_STYLE_TYPE.PARAGRAPH else "character",
            "font": font_info,
            "paragraph_format": para_info,
            "base_style": s.base_style.name if s.base_style else None,
        }

    def update_style(
        self,
        style_name: str,
        font_name: Optional[str] = None,
        font_size_pt: Optional[float] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        color_rgb: Optional[str] = None,
        alignment: Optional[str] = None,
        line_spacing: Optional[float] = None,
        space_before_pt: Optional[float] = None,
        space_after_pt: Optional[float] = None,
        first_line_indent_cm: Optional[float] = None,
    ) -> bool:
        """Updates properties of an existing style or creates it if not found."""
        try:
            s = self.model.doc.styles[style_name]
        except KeyError:
            # Create style
            s = self.model.doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)

        if font_name is not None:
            s.font.name = font_name
        if font_size_pt is not None:
            s.font.size = Pt(font_size_pt)
        if bold is not None:
            s.font.bold = bold
        if italic is not None:
            s.font.italic = italic
        if color_rgb is not None:
            s.font.color.rgb = hex_to_rgb(color_rgb)

        if s.type == WD_STYLE_TYPE.PARAGRAPH:
            fmt = s.paragraph_format
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

        return True

    def apply_style(
        self,
        target: Union[str, int, Dict[str, Any]],
        style_name: str,
    ) -> int:
        """Applies a style to target paragraph(s)."""
        paragraphs = self.resolver.resolve_paragraphs(target)
        for p in paragraphs:
            p.style = style_name
        return len(paragraphs)
