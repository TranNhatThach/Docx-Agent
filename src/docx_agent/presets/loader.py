"""
Preset loader and validator for formatting presets.
"""

import json
import os
from pathlib import Path
from typing import Dict, Any, Optional
from pydantic import BaseModel, Field
from docx_agent.core.exceptions import DocxAgentError, ErrorCode
from docx_agent.utils.paths import resolve_safe_path


class MarginsSchema(BaseModel):
    top_cm: float = 2.0
    bottom_cm: float = 2.0
    left_cm: float = 3.0
    right_cm: float = 2.0


class FontSchema(BaseModel):
    name: str = "Times New Roman"
    size_pt: float = 13.0
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    color_rgb: Optional[str] = None


class ParagraphFormatSchema(BaseModel):
    alignment: str = "justify"
    line_spacing: float = 1.5
    space_before_pt: float = 0.0
    space_after_pt: float = 6.0
    first_line_indent_cm: Optional[float] = 1.27
    keep_with_next: Optional[bool] = None


class StyleDefinitionSchema(BaseModel):
    font: FontSchema
    paragraph_format: Optional[ParagraphFormatSchema] = None


class PresetSchema(BaseModel):
    name: str
    description: str
    page_size: str = "a4"
    orientation: str = "portrait"
    margins: MarginsSchema = Field(default_factory=MarginsSchema)
    styles: Dict[str, StyleDefinitionSchema] = Field(default_factory=dict)
    header: Optional[Dict[str, Any]] = None
    footer: Optional[Dict[str, Any]] = None


PRESET_DIR = Path(__file__).parent


def load_preset(name_or_path: str) -> PresetSchema:
    """
    Loads a preset by built-in name (e.g. 'academic-vn', 'ieee', 'apa')
    or from a custom JSON file path.
    """
    clean_name = name_or_path.lower().replace("-", "_")
    builtin_path = PRESET_DIR / f"{clean_name}.json"

    if builtin_path.exists():
        target_path = builtin_path
    else:
        target_path = resolve_safe_path(name_or_path)

    if not target_path.exists():
        raise DocxAgentError(
            message=f"Formatting preset not found: '{name_or_path}'",
            code=ErrorCode.PRESET_NOT_FOUND,
            details={"requested": name_or_path, "resolved_path": str(target_path)},
            suggestion="Use one of built-in presets: 'academic-vn', 'ieee', 'apa', 'technical-report' or supply a valid JSON path.",
        )

    try:
        with open(target_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return PresetSchema(**data)
    except Exception as e:
        raise DocxAgentError(
            message=f"Failed to parse preset '{name_or_path}': {str(e)}",
            code=ErrorCode.PRESET_NOT_FOUND,
            details={"error": str(e)},
        )
