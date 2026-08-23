"""
Academic Document Engine: Applies institutional presets, typography hierarchies,
and thesis standards to documents.
"""

from typing import Optional, Union, Dict, Any, List
from docx_agent.core.document import DocumentModel
from docx_agent.presets.loader import load_preset, PresetSchema
from docx_agent.operations.sections import SectionOperations
from docx_agent.operations.styles import StyleOperations
from docx_agent.operations.footers import FooterOperations
from docx_agent.operations.headers import HeaderOperations


class AcademicOperations:
    """
    Applies unified institutional presets to documents.
    """

    def __init__(self, model: DocumentModel):
        self.model = model
        self.sec_ops = SectionOperations(model)
        self.style_ops = StyleOperations(model)
        self.footer_ops = FooterOperations(model)
        self.header_ops = HeaderOperations(model)

    def apply_preset(self, preset_name_or_path: str = "academic-vn") -> Dict[str, Any]:
        """
        Applies a full formatting preset to the document:
        - Configures page size and margins
        - Updates style definitions (Normal, Heading 1, Heading 2, Heading 3...)
        - Normalizes paragraph fonts and alignments where styles are applied
        - Configures header and footer (dynamic page numbers)
        """
        preset = load_preset(preset_name_or_path)

        # 1. Page Geometry & Margins
        self.sec_ops.configure_section(
            target=None,
            page_size=preset.page_size,
            orientation=preset.orientation,
            margin_top_cm=preset.margins.top_cm,
            margin_bottom_cm=preset.margins.bottom_cm,
            margin_left_cm=preset.margins.left_cm,
            margin_right_cm=preset.margins.right_cm,
        )

        # 2. Styles
        for style_name, style_def in preset.styles.items():
            font_kwargs = {
                "font_name": style_def.font.name,
                "font_size_pt": style_def.font.size_pt,
                "bold": style_def.font.bold,
                "italic": style_def.font.italic,
                "color_rgb": style_def.font.color_rgb,
            }
            para_kwargs = {}
            if style_def.paragraph_format:
                para_kwargs = {
                    "alignment": style_def.paragraph_format.alignment,
                    "line_spacing": style_def.paragraph_format.line_spacing,
                    "space_before_pt": style_def.paragraph_format.space_before_pt,
                    "space_after_pt": style_def.paragraph_format.space_after_pt,
                    "first_line_indent_cm": style_def.paragraph_format.first_line_indent_cm,
                }

            self.style_ops.update_style(
                style_name=style_name,
                **font_kwargs,
                **para_kwargs,
            )

        # 3. Footer / Header if defined in preset
        if preset.footer and preset.footer.get("page_numbers"):
            fmt = preset.footer.get("format", "Page {PAGE} of {NUMPAGES}")
            align = preset.footer.get("alignment", "center")
            self.footer_ops.set_page_numbers(target_section=None, format_str=fmt, alignment=align)

        self.model.reindex()

        return {
            "preset_applied": preset.name,
            "description": preset.description,
            "page_size": preset.page_size,
            "margins_cm": preset.margins.model_dump(),
            "styles_updated": list(preset.styles.keys()),
        }
