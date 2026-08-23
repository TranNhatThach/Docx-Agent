"""
Core subsystem exports for docx-agent.
"""

from docx_agent.core.document import DocumentModel
from docx_agent.core.identity import IdentityManager
from docx_agent.core.resolver import TargetResolver
from docx_agent.core.elements import (
    ElementType,
    ParagraphInfo,
    RunInfo,
    TableInfo,
    CellInfo,
    SectionInfo,
    DocumentSummary,
)
from docx_agent.core.exceptions import (
    ErrorCode,
    DocxAgentError,
    DocumentNotFoundError,
    InvalidDocxError,
    ElementNotFoundError,
    AmbiguousTargetError,
    StyleError,
    VerificationError,
    TransactionError,
)

__all__ = [
    "DocumentModel",
    "IdentityManager",
    "TargetResolver",
    "ElementType",
    "ParagraphInfo",
    "RunInfo",
    "TableInfo",
    "CellInfo",
    "SectionInfo",
    "DocumentSummary",
    "ErrorCode",
    "DocxAgentError",
    "DocumentNotFoundError",
    "InvalidDocxError",
    "ElementNotFoundError",
    "AmbiguousTargetError",
    "StyleError",
    "VerificationError",
    "TransactionError",
]
