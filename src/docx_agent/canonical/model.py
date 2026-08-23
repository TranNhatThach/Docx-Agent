"""
Canonical Document Model: Decoupled, high-performance in-memory runtime representation
for the AI-Native Document Workspace (Word-Grade OOXML Fidelity).
"""

from enum import Enum
from typing import Dict, Any, List, Optional, Union
from datetime import datetime
import uuid
from pydantic import BaseModel, Field


def generate_id(prefix: str = "node") -> str:
    """Generates a short, collision-resistant deterministic node identifier."""
    return f"{prefix}_{uuid.uuid4().hex[:8]}"


class DocumentProfile(str, Enum):
    ACADEMIC_REPORT = "academic_report"
    THESIS = "thesis"
    TECHNICAL_REPORT = "technical_report"
    RESEARCH_PAPER = "research_paper"
    BUSINESS_REPORT = "business_report"
    GENERAL = "general"


class BlockType(str, Enum):
    PARAGRAPH = "paragraph"
    HEADING = "heading"
    LIST_ITEM = "list_item"
    TABLE = "table"
    IMAGE = "image"
    DIAGRAM = "diagram"
    CITATION = "citation"
    UNSUPPORTED = "unsupported"


class ProvenanceType(str, Enum):
    HUMAN = "HUMAN"
    AGENT = "AGENT"
    IMPORT = "IMPORT"
    RESEARCH = "RESEARCH"


class ProvenanceRecord(BaseModel):
    source_type: ProvenanceType = ProvenanceType.HUMAN
    creator: str = "user"
    timestamp: str = Field(default_factory=lambda: datetime.now().isoformat())
    agent_transaction_id: Optional[str] = None
    source_url: Optional[str] = None
    doi: Optional[str] = None
    prompt: Optional[str] = None
    notes: Optional[str] = None


class RunNode(BaseModel):
    id: str = Field(default_factory=lambda: generate_id("r"))
    text: str = ""
    font_name: Optional[str] = None
    font_ascii: Optional[str] = None
    font_hAnsi: Optional[str] = None
    font_cs: Optional[str] = None
    font_eastAsia: Optional[str] = None
    font_size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    underline_style: Optional[str] = None
    strike: Optional[bool] = None
    dstrike: Optional[bool] = None
    color_rgb: Optional[str] = None
    highlight: Optional[str] = None
    superscript: Optional[bool] = None
    subscript: Optional[bool] = None
    hyperlink_url: Optional[str] = None
    citation_id: Optional[str] = None
    is_page_break: bool = False
    effective_formatting: Optional[Dict[str, Any]] = None


class BaseBlockNode(BaseModel):
    id: str = Field(default_factory=lambda: generate_id("blk"))
    type: BlockType = BlockType.PARAGRAPH
    style_name: Optional[str] = "Normal"
    dirty: bool = False
    provenance: ProvenanceRecord = Field(default_factory=ProvenanceRecord)


class ParagraphBlock(BaseBlockNode):
    type: BlockType = BlockType.PARAGRAPH
    runs: List[RunNode] = Field(default_factory=list)
    alignment: Optional[str] = "justify"
    line_spacing: Optional[float] = 1.5
    line_spacing_type: Optional[str] = "multiple"
    space_before_pt: Optional[float] = 0.0
    space_after_pt: Optional[float] = 6.0
    first_line_indent_cm: Optional[float] = 1.27
    hanging_indent_cm: Optional[float] = 0.0
    left_indent_cm: Optional[float] = 0.0
    right_indent_cm: Optional[float] = 0.0
    keep_with_next: Optional[bool] = False
    keep_lines: Optional[bool] = False
    page_break_before: Optional[bool] = False
    widow_control: Optional[bool] = True
    num_id: Optional[int] = None
    ilvl: Optional[int] = None
    resolved_numbering_label: Optional[str] = None
    effective_properties: Optional[Dict[str, Any]] = None

    @property
    def full_text(self) -> str:
        return "".join(r.text for r in self.runs)


class HeadingBlock(ParagraphBlock):
    type: BlockType = BlockType.HEADING
    level: int = 1
    style_name: Optional[str] = "Heading 1"
    first_line_indent_cm: Optional[float] = 0.0
    keep_with_next: Optional[bool] = True


class ListItemBlock(ParagraphBlock):
    type: BlockType = BlockType.LIST_ITEM
    list_type: str = "bullet"  # "bullet" or "number"
    list_level: int = 0
    first_line_indent_cm: Optional[float] = 0.0
    left_indent_cm: Optional[float] = 0.63


class TableCellNode(BaseModel):
    id: str = Field(default_factory=lambda: generate_id("cell"))
    text: str = ""
    runs: List[RunNode] = Field(default_factory=list)
    paragraphs: List[Any] = Field(default_factory=list)
    bg_color_hex: Optional[str] = None
    rowspan: int = 1
    colspan: int = 1
    width_cm: Optional[float] = None
    borders: Dict[str, Any] = Field(default_factory=dict)
    vertical_align: Optional[str] = "top"
    margin_top_cm: Optional[float] = None
    margin_bottom_cm: Optional[float] = None
    margin_left_cm: Optional[float] = None
    margin_right_cm: Optional[float] = None


class TableBlock(BaseBlockNode):
    type: BlockType = BlockType.TABLE
    rows: int = 0
    columns: int = 0
    cells: List[List[TableCellNode]] = Field(default_factory=list)
    style_name: Optional[str] = "Table Grid"
    alignment: str = "center"
    repeat_header: bool = True
    col_widths_cm: Optional[List[float]] = None
    tbl_width_cm: Optional[float] = None
    grid_cols_cm: List[float] = Field(default_factory=list)
    borders: Dict[str, Any] = Field(default_factory=dict)
    cant_split_rows: List[int] = Field(default_factory=list)


class ImageBlock(BaseBlockNode):
    type: BlockType = BlockType.IMAGE
    image_id: str = Field(default_factory=lambda: generate_id("img"))
    source_uri_or_path: str = ""
    source_kind: str = "LOCAL"  # "LOCAL", "WEB", "GENERATED", "DIAGRAM"
    width_cm: Optional[float] = 10.0
    height_cm: Optional[float] = None
    aspect_ratio: Optional[float] = None
    alignment: str = "center"
    position_mode: str = "inline"  # "inline" or "anchor"
    wrap_type: str = "inline"
    caption: Optional[str] = None


class DiagramBlock(BaseBlockNode):
    type: BlockType = BlockType.DIAGRAM
    diagram_id: str = Field(default_factory=lambda: generate_id("diag"))
    diagram_type: str = "architecture"  # "architecture", "flowchart", "sequence", "use_case", "er"
    source_code: str = ""  # Mermaid code or structured diagram definition
    rendered_svg: Optional[str] = None
    width_cm: Optional[float] = 12.0
    caption: Optional[str] = None


class UnsupportedBlock(BaseBlockNode):
    type: BlockType = BlockType.UNSUPPORTED
    tag_name: str = "unknown"
    namespace: Optional[str] = None
    attributes: Dict[str, str] = Field(default_factory=dict)
    raw_xml: str = ""
    original_order: int = 0
    warning_message: str = "Unsupported OOXML structure preserved without loss."


# Polymorphic block union
BlockNode = Union[
    HeadingBlock,
    ListItemBlock,
    ParagraphBlock,
    TableBlock,
    ImageBlock,
    DiagramBlock,
    UnsupportedBlock,
]


class SectionProperties(BaseModel):
    page_size: str = "a4"
    orientation: str = "portrait"
    page_width_cm: float = 21.0
    page_height_cm: float = 29.7
    margin_top_cm: float = 2.0
    margin_bottom_cm: float = 2.0
    margin_left_cm: float = 3.0
    margin_right_cm: float = 2.0
    columns: int = 1
    header_distance_cm: float = 1.27
    footer_distance_cm: float = 1.27
    different_first_page: bool = False


class SectionNode(BaseModel):
    id: str = Field(default_factory=lambda: generate_id("sec"))
    properties: SectionProperties = Field(default_factory=SectionProperties)
    header_text: Optional[str] = None
    footer_text: Optional[str] = None
    first_page_header_text: Optional[str] = None
    first_page_footer_text: Optional[str] = None
    has_page_numbers: bool = True
    page_number_format: str = "Page {PAGE} of {NUMPAGES}"
    page_numbering_start: int = 1
    blocks: List[BlockNode] = Field(default_factory=list)


class SourceMetadata(BaseModel):
    id: str = Field(default_factory=lambda: generate_id("src"))
    title: str
    authors: List[str] = Field(default_factory=list)
    publication: Optional[str] = None
    year: Optional[int] = None
    doi: Optional[str] = None
    url: Optional[str] = None
    source_type: str = "academic_paper"  # "academic_paper", "official_doc", "book", "web"
    retrieval_date: str = Field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d"))
    verified: bool = True
    confidence: float = 1.0


class CitationNode(BaseModel):
    id: str = Field(default_factory=lambda: generate_id("cite"))
    source_id: str
    citation_style: str = "apa"  # "apa", "ieee", "academic-vn"
    locator: Optional[str] = None  # e.g. "p. 42"
    formatted_intext: str = ""
    formatted_bibliography: str = ""
    provenance: ProvenanceRecord = Field(default_factory=ProvenanceRecord)


class DocumentNode(BaseModel):
    id: str = Field(default_factory=lambda: generate_id("doc"))
    title: str = "Untitled Document"
    profile: DocumentProfile = DocumentProfile.ACADEMIC_REPORT
    version: int = 1
    created_at: str = Field(default_factory=lambda: datetime.now().isoformat())
    updated_at: str = Field(default_factory=lambda: datetime.now().isoformat())
    sections: List[SectionNode] = Field(default_factory=lambda: [SectionNode()])
    sources: Dict[str, SourceMetadata] = Field(default_factory=dict)
    citations: Dict[str, CitationNode] = Field(default_factory=dict)
    custom_styles: Dict[str, Dict[str, Any]] = Field(default_factory=dict)
    raw_styles_xml: Optional[str] = None
    raw_numbering_xml: Optional[str] = None
    unknown_parts: Dict[str, str] = Field(default_factory=dict)

    def find_block(self, block_id: str) -> Optional[BaseBlockNode]:
        """Finds a block across all sections by its unique block_id."""
        for sec in self.sections:
            for blk in sec.blocks:
                if blk.id == block_id:
                    return blk
        return None

    def all_blocks(self) -> List[BaseBlockNode]:
        """Flattens all blocks across sections into a single list."""
        res = []
        for sec in self.sections:
            res.extend(sec.blocks)
        return res
