"""
Structured error taxonomy and exceptions for docx-agent.
Provides machine-readable error codes and rich diagnostics for AI agents.
"""

from enum import Enum
from typing import Any, Dict, Optional


class ErrorCode(str, Enum):
    """Stable error codes for machine parsing and agent self-repair."""
    FILE_NOT_FOUND = "FILE_NOT_FOUND"
    INVALID_DOCX = "INVALID_DOCX"
    ELEMENT_NOT_FOUND = "ELEMENT_NOT_FOUND"
    AMBIGUOUS_TARGET = "AMBIGUOUS_TARGET"
    INVALID_OPERATION = "INVALID_OPERATION"
    INVALID_STYLE = "INVALID_STYLE"
    INVALID_SECTION = "INVALID_SECTION"
    FORMAT_ERROR = "FORMAT_ERROR"
    TABLE_ERROR = "TABLE_ERROR"
    IMAGE_ERROR = "IMAGE_ERROR"
    OOXML_ERROR = "OOXML_ERROR"
    VERIFICATION_FAILED = "VERIFICATION_FAILED"
    TRANSACTION_FAILED = "TRANSACTION_FAILED"
    ROLLBACK_FAILED = "ROLLBACK_FAILED"
    PRESET_NOT_FOUND = "PRESET_NOT_FOUND"
    UNSUPPORTED_ELEMENT = "UNSUPPORTED_ELEMENT"


class DocxAgentError(Exception):
    """Base exception for all docx-agent errors."""

    def __init__(
        self,
        message: str,
        code: ErrorCode = ErrorCode.INVALID_OPERATION,
        details: Optional[Dict[str, Any]] = None,
        suggestion: Optional[str] = None,
    ):
        super().__init__(message)
        self.message = message
        self.code = code
        self.details = details or {}
        self.suggestion = suggestion

    def to_dict(self) -> Dict[str, Any]:
        """Returns structured error payload for JSON reporting."""
        payload: Dict[str, Any] = {
            "code": self.code.value,
            "message": self.message,
        }
        if self.details:
            payload["details"] = self.details
        if self.suggestion:
            payload["suggestion"] = self.suggestion
        return payload


class DocumentNotFoundError(DocxAgentError):
    def __init__(self, path: str):
        super().__init__(
            message=f"File not found: {path}",
            code=ErrorCode.FILE_NOT_FOUND,
            details={"path": path},
            suggestion="Verify file path and existence before invoking operations."
        )


class InvalidDocxError(DocxAgentError):
    def __init__(self, path: str, reason: str):
        super().__init__(
            message=f"Invalid or corrupted DOCX file: {path}. Reason: {reason}",
            code=ErrorCode.INVALID_DOCX,
            details={"path": path, "reason": reason},
            suggestion="Ensure the file is a valid Microsoft Word .docx OpenXML package."
        )


class ElementNotFoundError(DocxAgentError):
    def __init__(self, target: str, target_type: str = "element"):
        super().__init__(
            message=f"{target_type.capitalize()} not found matching target: '{target}'",
            code=ErrorCode.ELEMENT_NOT_FOUND,
            details={"target": target, "target_type": target_type},
            suggestion="Use 'docx-agent inspect' or 'docx-agent find' to locate valid element IDs or selectors."
        )


class AmbiguousTargetError(DocxAgentError):
    def __init__(self, target: str, match_count: int, matched_ids: list):
        super().__init__(
            message=f"Target '{target}' matched {match_count} elements. Expected exactly 1 match.",
            code=ErrorCode.AMBIGUOUS_TARGET,
            details={"target": target, "match_count": match_count, "matched_ids": matched_ids},
            suggestion="Specify exact element_id (e.g. 'p_0012') or refine the selector to disambiguate."
        )


class StyleError(DocxAgentError):
    def __init__(self, style_name: str, message: str):
        super().__init__(
            message=f"Style error for '{style_name}': {message}",
            code=ErrorCode.INVALID_STYLE,
            details={"style_name": style_name, "error": message},
            suggestion="Use 'docx-agent styles' to inspect available styles in the document."
        )


class VerificationError(DocxAgentError):
    def __init__(self, message: str, failures: list):
        super().__init__(
            message=message,
            code=ErrorCode.VERIFICATION_FAILED,
            details={"failures": failures},
            suggestion="Review verification failures and apply corrective formatting or structure operations."
        )


class TransactionError(DocxAgentError):
    def __init__(self, message: str, step: str, original_error: Optional[Exception] = None):
        super().__init__(
            message=f"Transaction failed during step '{step}': {message}",
            code=ErrorCode.TRANSACTION_FAILED,
            details={"step": step, "original_error": str(original_error) if original_error else None},
            suggestion="Check document write permissions and ensure no other process has locked the file."
        )


class InvalidOperationError(DocxAgentError):
    def __init__(self, message: str, operation_name: str = "operation"):
        super().__init__(
            message=f"Invalid operation '{operation_name}': {message}",
            code=ErrorCode.INVALID_OPERATION,
            details={"operation": operation_name},
            suggestion="Verify operation arguments and target element type before execution."
        )


class FormatError(DocxAgentError):
    def __init__(self, property_name: str, value: Any, message: str):
        super().__init__(
            message=f"Formatting error for property '{property_name}' with value '{value}': {message}",
            code=ErrorCode.FORMAT_ERROR,
            details={"property": property_name, "value": str(value), "error": message},
            suggestion="Ensure formatting value is within supported type and range constraints."
        )


class TableError(DocxAgentError):
    def __init__(self, message: str, table_id: Optional[str] = None):
        super().__init__(
            message=f"Table error{' for ' + table_id if table_id else ''}: {message}",
            code=ErrorCode.TABLE_ERROR,
            details={"table_id": table_id},
            suggestion="Check table dimensions, cell boundaries, and grid spans."
        )


class ImageError(DocxAgentError):
    def __init__(self, message: str, image_path: Optional[str] = None):
        super().__init__(
            message=f"Image processing error{' for ' + image_path if image_path else ''}: {message}",
            code=ErrorCode.IMAGE_ERROR,
            details={"image_path": image_path},
            suggestion="Ensure the image exists, is in supported format (PNG/JPEG/SVG), and has valid dimensions."
        )


class RollbackFailedError(DocxAgentError):
    def __init__(self, message: str, backup_path: Optional[str] = None):
        super().__init__(
            message=f"Transaction rollback failed: {message}",
            code=ErrorCode.ROLLBACK_FAILED,
            details={"backup_path": backup_path},
            suggestion="Manual recovery from backup file may be required."
        )


class PresetNotFoundError(DocxAgentError):
    def __init__(self, preset_name: str):
        super().__init__(
            message=f"Document preset profile not found: '{preset_name}'",
            code=ErrorCode.PRESET_NOT_FOUND,
            details={"preset_name": preset_name},
            suggestion="Check available presets using 'docx-agent presets list'."
        )


class UnsupportedElementError(DocxAgentError):
    def __init__(self, element_tag: str, message: str = "Element is preserved in unsupported raw model."):
        super().__init__(
            message=f"Unsupported OOXML element <{element_tag}>: {message}",
            code=ErrorCode.UNSUPPORTED_ELEMENT,
            details={"tag": element_tag},
            suggestion="Element will be preserved intact during roundtrip export."
        )

