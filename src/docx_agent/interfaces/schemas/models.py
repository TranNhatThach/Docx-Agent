"""
Pydantic schemas and DTOs for agent interfaces, plans, and machine-readable responses.
"""

from typing import List, Dict, Any, Optional, Union
from pydantic import BaseModel, Field


class PlanOperationSchema(BaseModel):
    type: str  # "replace", "insert", "delete", "append", "format_text", "format_para", "style", "table", "preset", "page_number", "header", "footer", "toc"
    target: Optional[Union[str, int, Dict[str, Any]]] = None
    text: Optional[str] = None
    replacement: Optional[str] = None
    position: Optional[str] = "after"
    style: Optional[str] = None
    level: Optional[int] = None
    font_name: Optional[str] = None
    font_size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    color_rgb: Optional[str] = None
    alignment: Optional[str] = None
    line_spacing: Optional[float] = None
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None
    first_line_indent_cm: Optional[float] = None
    preset: Optional[str] = None
    format: Optional[str] = None
    rows: Optional[int] = None
    cols: Optional[int] = None
    data: Optional[List[List[str]]] = None
    image_path: Optional[str] = None
    width_cm: Optional[float] = None
    caption: Optional[str] = None


class BatchPlanSchema(BaseModel):
    document: str
    operations: List[PlanOperationSchema] = Field(default_factory=list)
    verify: bool = True
    auto_backup: bool = True


class AgentResponse(BaseModel):
    success: bool
    operation: str
    document: Optional[str] = None
    matched: Optional[int] = None
    changed: Optional[int] = None
    data: Optional[Any] = None
    verification: Optional[Dict[str, Any]] = None
    error: Optional[Dict[str, Any]] = None
