"""
Universal Open-Source DOCX Agent Platform.
"""

from docx_agent.agent import DocumentAgent
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

__version__ = "1.0.0"

__all__ = [
    "DocumentAgent",
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
