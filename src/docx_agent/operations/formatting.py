"""
Text and character formatting operations.
"""

from typing import Optional, Union, Dict, Any, List
import docx
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_COLOR_INDEX
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx_agent.core.document import DocumentModel
from docx_agent.core.resolver import TargetResolver


HIGHLIGHT_MAP = {
    "yellow": WD_COLOR_INDEX.YELLOW,
    "green": WD_COLOR_INDEX.BRIGHT_GREEN,
    "cyan": WD_COLOR_INDEX.TURQUOISE,
    "magenta": WD_COLOR_INDEX.PINK,
    "blue": WD_COLOR_INDEX.BLUE,
    "red": WD_COLOR_INDEX.RED,
    "dark_blue": WD_COLOR_INDEX.DARK_BLUE,
    "dark_cyan": WD_COLOR_INDEX.TEAL,
    "dark_green": WD_COLOR_INDEX.GREEN,
    "dark_magenta": WD_COLOR_INDEX.VIOLET,
    "dark_red": WD_COLOR_INDEX.DARK_RED,
    "dark_yellow": WD_COLOR_INDEX.DARK_YELLOW,
    "gray_50": WD_COLOR_INDEX.GRAY_50,
    "gray_25": WD_COLOR_INDEX.GRAY_25,
    "black": WD_COLOR_INDEX.BLACK,
}


def hex_to_rgb(hex_str: str) -> RGBColor:
    """Parses '#RRGGBB' or 'RRGGBB' into docx.shared.RGBColor."""
    clean = hex_str.lstrip("#")
    if len(clean) != 6:
        raise ValueError(f"Invalid hex color string: {hex_str}")
    r = int(clean[0:2], 16)
    g = int(clean[2:4], 16)
    b = int(clean[4:6], 16)
    return RGBColor(r, g, b)


class TextFormattingOperations:
    """
    Applies font and character formatting to paragraphs, runs, or text selections.
    """

    def __init__(self, model: DocumentModel):
        self.model = model
        self.resolver = TargetResolver(model)

    def format_text(
        self,
        target: Union[str, int, Dict[str, Any]],
        font_name: Optional[str] = None,
        font_size_pt: Optional[float] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        underline: Optional[bool] = None,
        strike: Optional[bool] = None,
        color_rgb: Optional[str] = None,
        highlight: Optional[str] = None,
        superscript: Optional[bool] = None,
        subscript: Optional[bool] = None,
    ) -> int:
        """
        Formats text runs in the matched paragraph(s).
        Returns count of paragraphs formatted.
        """
        paragraphs = self.resolver.resolve_paragraphs(target)
        for p in paragraphs:
            for r in p.runs:
                if font_name is not None:
                    r.font.name = font_name
                if font_size_pt is not None:
                    r.font.size = Pt(font_size_pt)
                if bold is not None:
                    r.bold = bold
                if italic is not None:
                    r.italic = italic
                if underline is not None:
                    r.underline = underline
                if strike is not None:
                    r.font.strike = strike
                if color_rgb is not None:
                    r.font.color.rgb = hex_to_rgb(color_rgb)
                if highlight is not None:
                    h_val = HIGHLIGHT_MAP.get(highlight.lower())
                    if h_val:
                        r.font.highlight_color = h_val
                if superscript is not None:
                    r.font.superscript = superscript
                if subscript is not None:
                    r.font.subscript = subscript

        return len(paragraphs)
