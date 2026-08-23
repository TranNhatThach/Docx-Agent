"""
Semantic Change Model: Core data structures for AI Change Tracking, Review, and Version Control.
Independent of raw DOCX XML, anchored stably to canonical document elements.
"""

from enum import Enum
from typing import Dict, Any, List, Optional, Union
from datetime import datetime
import uuid
import hashlib
from pydantic import BaseModel, Field


def generate_change_id(prefix: str = "CHG") -> str:
    """Generates a human-friendly change identifier, e.g., CHG-001 or CHG-a1b2c3d4."""
    return f"{prefix}-{uuid.uuid4().hex[:6].upper()}"


def generate_session_id(prefix: str = "SES") -> str:
    """Generates a human-friendly edit session identifier, e.g., SES-001 or SES-a1b2c3d4."""
    return f"{prefix}-{uuid.uuid4().hex[:6].upper()}"


class ChangeType(str, Enum):
    # Text-level changes
    INSERT_TEXT = "INSERT_TEXT"
    DELETE_TEXT = "DELETE_TEXT"
    REPLACE_TEXT = "REPLACE_TEXT"
    REORDER_TEXT = "REORDER_TEXT"

    # Paragraph-level changes
    ADD_PARAGRAPH = "ADD_PARAGRAPH"
    DELETE_PARAGRAPH = "DELETE_PARAGRAPH"
    MODIFY_PARAGRAPH = "MODIFY_PARAGRAPH"
    MOVE_PARAGRAPH = "MOVE_PARAGRAPH"

    # Heading-level changes
    ADD_HEADING = "ADD_HEADING"
    REMOVE_HEADING = "REMOVE_HEADING"
    CHANGE_HEADING_LEVEL = "CHANGE_HEADING_LEVEL"
    MODIFY_HEADING_TEXT = "MODIFY_HEADING_TEXT"

    # Table-level changes
    ADD_TABLE = "ADD_TABLE"
    DELETE_TABLE = "DELETE_TABLE"
    ADD_ROW = "ADD_ROW"
    DELETE_ROW = "DELETE_ROW"
    ADD_COLUMN = "ADD_COLUMN"
    DELETE_COLUMN = "DELETE_COLUMN"
    MODIFY_CELL_CONTENT = "MODIFY_CELL_CONTENT"
    MODIFY_CELL_FORMATTING = "MODIFY_CELL_FORMATTING"

    # Formatting changes
    FORMAT_TEXT = "FORMAT_TEXT"
    FORMAT_PARAGRAPH = "FORMAT_PARAGRAPH"
    CHANGE_STYLE = "CHANGE_STYLE"

    # Structural changes
    ADD_PAGE_BREAK = "ADD_PAGE_BREAK"
    DELETE_PAGE_BREAK = "DELETE_PAGE_BREAK"
    CHANGE_SECTION = "CHANGE_SECTION"
    MODIFY_HEADER = "MODIFY_HEADER"
    MODIFY_FOOTER = "MODIFY_FOOTER"
    ADD_IMAGE = "ADD_IMAGE"
    DELETE_IMAGE = "DELETE_IMAGE"
    ADD_CITATION = "ADD_CITATION"
    MODIFY_NUMBERING = "MODIFY_NUMBERING"
    COMPOSITE = "COMPOSITE"


class ChangeStatus(str, Enum):
    PROPOSED = "PROPOSED"
    ACCEPTED = "ACCEPTED"
    REJECTED = "REJECTED"
    REVERTED = "REVERTED"


class ChangeLocation(BaseModel):
    """
    Stable semantic anchor for document modifications.
    Does NOT rely strictly on volatile paragraph indices.
    """
    document_id: Optional[str] = None
    section_id: Optional[str] = None
    block_id: Optional[str] = None
    paragraph_id: Optional[str] = None
    table_id: Optional[str] = None
    row_idx: Optional[int] = None
    col_idx: Optional[int] = None
    cell_id: Optional[str] = None
    run_id: Optional[str] = None
    start_offset: int = 0
    end_offset: int = 0
    text_anchor: Optional[str] = None  # Snippet of surrounding context before mutation
    semantic_hash: Optional[str] = None  # Content hash of the target container

    @staticmethod
    def compute_semantic_hash(content: str) -> str:
        """Computes short MD5 hash for semantic content verification."""
        return hashlib.md5(content.strip().encode("utf-8")).hexdigest()[:12]


class ChangeObject(BaseModel):
    """
    Canonical Change Representation for every AI or human modification.
    Enables granular Explainability, Review, Accept/Reject, Undo, and Audit.
    """
    change_id: str = Field(default_factory=generate_change_id)
    session_id: Optional[str] = None
    document_id: str = "doc_default"
    version_before: int = 1
    version_after: int = 2
    change_type: ChangeType = ChangeType.MODIFY_PARAGRAPH
    target_element: str = "unknown"  # e.g., p_0012, tbl_0001, h_0003
    location: ChangeLocation = Field(default_factory=ChangeLocation)
    before_content: Optional[Any] = None
    after_content: Optional[Any] = None
    reason: str = "Tối ưu hóa nội dung văn bản theo ngữ cảnh học thuật."
    agent_id: str = "Antigravity-Agent"
    agent_action: str = "improve_clarity"
    timestamp: str = Field(default_factory=lambda: datetime.now().isoformat())
    status: ChangeStatus = ChangeStatus.PROPOSED
    confidence: float = Field(default=0.95, ge=0.0, le=1.0)
    evidence: str = "Evidence: NOT AVAILABLE"
    affected_formatting: Dict[str, Any] = Field(default_factory=dict)
    affected_structure: Dict[str, Any] = Field(default_factory=dict)
    reverted_at: Optional[str] = None
    accepted_at: Optional[str] = None
    rejected_at: Optional[str] = None

    def accept(self) -> None:
        """Marks this change as accepted by the user."""
        self.status = ChangeStatus.ACCEPTED
        self.accepted_at = datetime.now().isoformat()

    def reject(self) -> None:
        """Marks this change as rejected by the user."""
        self.status = ChangeStatus.REJECTED
        self.rejected_at = datetime.now().isoformat()

    def revert(self) -> None:
        """Marks this change as reverted after prior acceptance."""
        self.status = ChangeStatus.REVERTED
        self.reverted_at = datetime.now().isoformat()

    def is_pending(self) -> bool:
        return self.status == ChangeStatus.PROPOSED

    def is_accepted(self) -> bool:
        return self.status == ChangeStatus.ACCEPTED

    def is_rejected(self) -> bool:
        return self.status == ChangeStatus.REJECTED

    def is_reverted(self) -> bool:
        return self.status == ChangeStatus.REVERTED


class EditSession(BaseModel):
    """
    Encapsulates a batch AI modification task with discrete, auditable Change Objects.
    No changes are committed to the main document until user review.
    """
    session_id: str = Field(default_factory=generate_session_id)
    task_description: str = "Chỉnh sửa và cải thiện tài liệu"
    agent_id: str = "Antigravity-Agent"
    document_id: str = "doc_default"
    base_version: int = 1
    target_version: int = 2
    created_at: str = Field(default_factory=lambda: datetime.now().isoformat())
    status: str = "OPEN"  # "OPEN", "REVIEWING", "COMMITTED", "DISCARDED"
    changes: List[ChangeObject] = Field(default_factory=list)
    metadata: Dict[str, Any] = Field(default_factory=dict)

    def add_change(self, change: ChangeObject) -> None:
        change.session_id = self.session_id
        change.document_id = self.document_id
        change.version_before = self.base_version
        change.version_after = self.target_version
        self.changes.append(change)

    def get_change(self, change_id: str) -> Optional[ChangeObject]:
        for chg in self.changes:
            if chg.change_id == change_id:
                return chg
        return None

    def accept_change(self, change_id: str) -> bool:
        chg = self.get_change(change_id)
        if chg:
            chg.accept()
            return True
        return False

    def reject_change(self, change_id: str) -> bool:
        chg = self.get_change(change_id)
        if chg:
            chg.reject()
            return True
        return False

    def accept_all(self) -> int:
        count = 0
        for chg in self.changes:
            if chg.status == ChangeStatus.PROPOSED:
                chg.accept()
                count += 1
        return count

    def reject_all(self) -> int:
        count = 0
        for chg in self.changes:
            if chg.status == ChangeStatus.PROPOSED:
                chg.reject()
                count += 1
        return count

    def pending_count(self) -> int:
        return sum(1 for c in self.changes if c.status == ChangeStatus.PROPOSED)

    def accepted_count(self) -> int:
        return sum(1 for c in self.changes if c.status == ChangeStatus.ACCEPTED)

    def rejected_count(self) -> int:
        return sum(1 for c in self.changes if c.status == ChangeStatus.REJECTED)

    def summary(self) -> Dict[str, Any]:
        return {
            "session_id": self.session_id,
            "task": self.task_description,
            "agent_id": self.agent_id,
            "base_version": self.base_version,
            "target_version": self.target_version,
            "total_changes": len(self.changes),
            "pending": self.pending_count(),
            "accepted": self.accepted_count(),
            "rejected": self.rejected_count(),
            "status": self.status,
        }
