"""
Element representations and metadata models for docx-agent.
"""

from enum import Enum
from typing import Any, Dict, List, Optional
from pydantic import BaseModel, Field


class ElementType(str, Enum):
    DOCUMENT = "document"
    SECTION = "section"
    PARAGRAPH = "paragraph"
    RUN = "run"
    TABLE = "table"
    ROW = "row"
    CELL = "cell"
    IMAGE = "image"
    HEADER = "header"
    FOOTER = "footer"


class RunInfo(BaseModel):
    id: str
    text: str
    font_name: Optional[str] = None
    font_size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    strike: Optional[bool] = None
    color_rgb: Optional[str] = None
    highlight: Optional[str] = None
    superscript: Optional[bool] = None
    subscript: Optional[bool] = None


class ParagraphInfo(BaseModel):
    id: str
    type: ElementType = ElementType.PARAGRAPH
    index: int
    text: str
    style: str
    is_heading: bool = False
    heading_level: Optional[int] = None
    alignment: Optional[str] = None
    runs_count: int = 0
    section_index: int = 0
    runs: Optional[List[RunInfo]] = None


class CellInfo(BaseModel):
    id: str
    row_idx: int
    col_idx: int
    text: str
    paragraphs_count: int = 1


class TableInfo(BaseModel):
    id: str
    type: ElementType = ElementType.TABLE
    index: int
    rows: int
    columns: int
    style: Optional[str] = None
    preview: List[List[str]] = Field(default_factory=list)


class SectionInfo(BaseModel):
    id: str
    type: ElementType = ElementType.SECTION
    index: int
    page_width_cm: float
    page_height_cm: float
    orientation: str
    margin_top_cm: float
    margin_bottom_cm: float
    margin_left_cm: float
    margin_right_cm: float
    has_header: bool = False
    has_footer: bool = False


class ImageInfo(BaseModel):
    id: str
    type: ElementType = ElementType.IMAGE
    paragraph_id: str
    rel_id: str
    filename: Optional[str] = None
    width_pt: Optional[float] = None
    height_pt: Optional[float] = None


class DocumentSummary(BaseModel):
    file_path: Optional[str] = None
    paragraphs_count: int
    tables_count: int
    sections_count: int
    images_count: int = 0
    styles: List[str] = Field(default_factory=list)
    headings_outline: List[Dict[str, Any]] = Field(default_factory=list)
    sections: List[SectionInfo] = Field(default_factory=list)
