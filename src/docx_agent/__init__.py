"""
Universal Open-Source DOCX Agent Platform & AI-Native Document Workspace.
"""

from docx_agent.agent import DocumentAgent
from docx_agent.canonical.model import (
    BlockNode,
    DiagramBlock,
    DocumentNode,
    HeadingBlock,
    ImageBlock,
    ListItemBlock,
    ParagraphBlock,
    RunNode,
    SectionNode,
    TableBlock,
    TableCellNode,
    UnsupportedBlock,
)
from docx_agent.core.config import Settings, get_config
from docx_agent.core.exceptions import (
    AmbiguousTargetError,
    DocxAgentError,
    DocumentNotFoundError,
    ElementNotFoundError,
    ErrorCode,
    FormatError,
    ImageError,
    InvalidDocxError,
    InvalidOperationError,
    PresetNotFoundError,
    RollbackFailedError,
    StyleError,
    TableError,
    TransactionError,
    UnsupportedElementError,
    VerificationError,
)
from docx_agent.adapters.docx import DocxExporter, DocxImporter
from docx_agent.operations.markdown import MarkdownToDocxConverter
from docx_agent.engine.layout import LayoutDocument, LayoutEngine, LayoutPage
from docx_agent.engine.numbering import NumberingResolver
from docx_agent.engine.styles import StyleResolver
from docx_agent.interfaces.workspace.bridge import WorkspaceBridge
from docx_agent.transactions.transaction import TransactionContext

__version__ = "2.1.0"

__all__ = [
    # Top-Level Agent
    "DocumentAgent",
    "__version__",
    # Configuration
    "Settings",
    "get_config",
    # Canonical Model
    "DocumentNode",
    "SectionNode",
    "BlockNode",
    "ParagraphBlock",
    "HeadingBlock",
    "ListItemBlock",
    "TableBlock",
    "TableCellNode",
    "ImageBlock",
    "DiagramBlock",
    "UnsupportedBlock",
    "RunNode",
    # Engine & Resolvers
    "LayoutEngine",
    "LayoutDocument",
    "LayoutPage",
    "StyleResolver",
    "NumberingResolver",
    # Adapters & I/O
    "DocxImporter",
    "DocxExporter",
    "MarkdownToDocxConverter",
    # Workspace & Transactions
    "WorkspaceBridge",
    "TransactionContext",
    # Exceptions & Error Taxonomy
    "ErrorCode",
    "DocxAgentError",
    "DocumentNotFoundError",
    "InvalidDocxError",
    "ElementNotFoundError",
    "AmbiguousTargetError",
    "InvalidOperationError",
    "StyleError",
    "FormatError",
    "TableError",
    "ImageError",
    "VerificationError",
    "TransactionError",
    "RollbackFailedError",
    "PresetNotFoundError",
    "UnsupportedElementError",
]
