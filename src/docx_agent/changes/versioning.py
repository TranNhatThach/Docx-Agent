"""
Version Management & Audit History: Immutable Document Versioning, Snapshots,
Change-level Undo/Redo, Reverting, and Inter-version Semantic Diffing.
"""

from typing import Dict, Any, List, Optional, Tuple
from datetime import datetime
from copy import deepcopy
from pydantic import BaseModel, Field

from docx_agent.canonical.model import DocumentNode
from docx_agent.changes.model import (
    ChangeObject,
    ChangeStatus,
    ChangeType,
    EditSession,
    generate_change_id,
)


class VersionSnapshot(BaseModel):
    """
    Immutable document snapshot representing a committed state in history.
    """
    version: int
    timestamp: str = Field(default_factory=lambda: datetime.now().isoformat())
    author: str = "AI-Agent"
    description: str = "Cập nhật tài liệu"
    session_id: Optional[str] = None
    changes_applied: List[str] = Field(default_factory=list)
    changes_rejected: List[str] = Field(default_factory=list)
    document_snapshot: DocumentNode


class VersionManager:
    """
    Manages complete document version history, atomic commits, rollbacks, and change-level undo.
    """

    def __init__(self, initial_doc: DocumentNode):
        self.history: Dict[int, VersionSnapshot] = {}
        self.current_version: int = initial_doc.version or 1
        self.all_changes: Dict[str, ChangeObject] = {}
        self.sessions: Dict[str, EditSession] = {}
        self.change_undo_stack: List[ChangeObject] = []
        self.change_redo_stack: List[ChangeObject] = []

        # Record Initial Version 1
        self.record_initial_version(initial_doc)

    def record_initial_version(self, doc: DocumentNode) -> VersionSnapshot:
        snap = VersionSnapshot(
            version=1,
            author="Human / Original",
            description="Tài liệu ban đầu (Original Base Version)",
            document_snapshot=deepcopy(doc),
        )
        self.history[1] = snap
        self.current_version = 1
        return snap

    def register_session(self, session: EditSession) -> None:
        self.sessions[session.session_id] = session
        for chg in session.changes:
            self.all_changes[chg.change_id] = chg

    def create_edit_session(
        self,
        task_description: str,
        modified_doc: DocumentNode,
        agent_id: str = "Antigravity-Agent",
        reason: Optional[str] = None,
        confidence: float = 0.95,
        evidence: Optional[str] = None,
    ) -> EditSession:
        """
        Creates a new EditSession with discrete semantic ChangeObjects computed against current version snapshot.
        """
        from docx_agent.changes.diff import SemanticDiffEngine
        current_snap = self.history.get(self.current_version)
        current_doc = current_snap.document_snapshot if current_snap else DocumentNode()
        diff_engine = SemanticDiffEngine()
        changes = diff_engine.compute_diff(
            current_doc,
            modified_doc,
            agent_id=agent_id,
            default_reason=reason or "Cập nhật và tối ưu hóa nội dung",
            confidence=confidence,
        )
        session = EditSession(
            task_description=task_description,
            base_version=self.current_version,
            agent_id=agent_id,
            changes=changes,
        )
        self.register_session(session)
        return session

    def commit_session(
        self,
        session_id: str,
        current_doc: Optional[DocumentNode] = None,
        author: str = "User & AI Agent",
    ) -> VersionSnapshot:
        session = self.sessions.get(session_id)
        if not session:
            raise ValueError(f"Không tìm thấy phiên chỉnh sửa (Edit Session): {session_id}")

        if current_doc is None:
            base_snap = self.history.get(self.current_version)
            current_doc = deepcopy(base_snap.document_snapshot) if base_snap else DocumentNode()

        accepted_ids = []
        rejected_ids = []
        for chg in session.changes:
            if chg.status == ChangeStatus.ACCEPTED:
                accepted_ids.append(chg.change_id)
                self.change_undo_stack.append(chg)
            elif chg.status == ChangeStatus.REJECTED:
                rejected_ids.append(chg.change_id)

        self.current_version += 1
        current_doc.version = self.current_version
        current_doc.updated_at = datetime.now().isoformat()
        session.status = "COMMITTED"
        session.target_version = self.current_version

        snapshot = VersionSnapshot(
            version=self.current_version,
            author=author,
            description=f"Commit [{session.session_id}]: {session.task_description} ({len(accepted_ids)} chấp nhận, {len(rejected_ids)} từ chối)",
            session_id=session.session_id,
            changes_applied=accepted_ids,
            changes_rejected=rejected_ids,
            document_snapshot=deepcopy(current_doc),
        )
        self.history[self.current_version] = snapshot
        return snapshot

    def get_version(self, version_num: int) -> Optional[VersionSnapshot]:
        return self.history.get(version_num)

    def restore_version(self, version_num: int) -> DocumentNode:
        snap = self.history.get(version_num)
        if not snap:
            raise ValueError(f"Không tìm thấy phiên bản {version_num} trong lịch sử.")

        self.current_version += 1
        restored = deepcopy(snap.document_snapshot)
        restored.version = self.current_version
        restored.updated_at = datetime.now().isoformat()

        # Record restoration as a new version
        new_snap = VersionSnapshot(
            version=self.current_version,
            author="User",
            description=f"Khôi phục về trạng thái Phiên bản {version_num}",
            document_snapshot=deepcopy(restored),
        )
        self.history[self.current_version] = new_snap
        return restored

    def revert_change(self, change_id: str, current_doc: DocumentNode) -> Tuple[DocumentNode, bool]:
        chg = self.all_changes.get(change_id)
        if not chg:
            return current_doc, False

        chg.revert()
        # Find target element and restore before_content if applicable
        blk = current_doc.find_block(chg.target_element)
        if blk and chg.before_content is not None:
            if hasattr(blk, "runs") and isinstance(chg.before_content, str):
                from docx_agent.canonical.model import RunNode
                blk.runs = [RunNode(text=chg.before_content)]
                blk.dirty = True

        self.current_version += 1
        current_doc.version = self.current_version

        new_snap = VersionSnapshot(
            version=self.current_version,
            author="User",
            description=f"Hoàn tác thay đổi {change_id} ({chg.change_type.value})",
            changes_rejected=[change_id],
            document_snapshot=deepcopy(current_doc),
        )
        self.history[self.current_version] = new_snap
        return current_doc, True

    def undo_last_change(self, current_doc: DocumentNode) -> Tuple[DocumentNode, Optional[ChangeObject]]:
        if not self.change_undo_stack:
            return current_doc, None

        chg = self.change_undo_stack.pop()
        current_doc, success = self.revert_change(chg.change_id, current_doc)
        if success:
            self.change_redo_stack.append(chg)
            return current_doc, chg
        return current_doc, None

    def redo_last_change(self, current_doc: DocumentNode) -> Tuple[DocumentNode, Optional[ChangeObject]]:
        if not self.change_redo_stack:
            return current_doc, None

        chg = self.change_redo_stack.pop()
        chg.accept()
        blk = current_doc.find_block(chg.target_element)
        if blk and chg.after_content is not None:
            if hasattr(blk, "runs") and isinstance(chg.after_content, str):
                from docx_agent.canonical.model import RunNode
                blk.runs = [RunNode(text=chg.after_content)]
                blk.dirty = True

        self.current_version += 1
        current_doc.version = self.current_version
        self.change_undo_stack.append(chg)
        return current_doc, chg

    def list_history(self) -> List[Dict[str, Any]]:
        res = []
        for v in sorted(self.history.keys()):
            snap = self.history[v]
            res.append({
                "version": snap.version,
                "timestamp": snap.timestamp,
                "author": snap.author,
                "description": snap.description,
                "session_id": snap.session_id,
                "changes_applied_count": len(snap.changes_applied),
                "changes_rejected_count": len(snap.changes_rejected),
            })
        return res

    def get_history(self) -> List[Dict[str, Any]]:
        return self.list_history()

