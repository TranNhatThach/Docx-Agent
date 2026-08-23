"""
Persistence, Operation Log, Periodic Snapshots, and Crash Recovery Engine.
"""

import json
import os
from pathlib import Path
from typing import Dict, Any, List, Optional
from datetime import datetime
from pydantic import BaseModel, Field

from docx_agent.canonical.model import DocumentNode
from docx_agent.engine.operations import DocOperation
from docx_agent.utils.paths import resolve_safe_path, ensure_parent_dir


class LogEntry(BaseModel):
    revision: int
    timestamp: str = Field(default_factory=lambda: datetime.now().isoformat())
    author: str = "human"  # "human" or "agent"
    op_type: str
    payload: Dict[str, Any]


class WorkspacePersistence:
    """
    Manages local workspace storage, sequential operation log, periodic snapshots,
    and automatic crash recovery.
    """

    def __init__(self, workspace_dir: Optional[Path] = None):
        self.workspace_dir = workspace_dir or Path(".docx_agent_workspace")
        self.snapshots_dir = self.workspace_dir / "snapshots"
        self.log_file = self.workspace_dir / "operation_log.jsonl"
        self._init_storage()

    def _init_storage(self) -> None:
        self.snapshots_dir.mkdir(parents=True, exist_ok=True)

    def record_operation(self, revision: int, op: DocOperation, author: str = "human") -> None:
        """Appends an executed operation to the persistent operation log."""
        entry = LogEntry(
            revision=revision,
            author=author,
            op_type=op.op_type,
            payload=op.model_dump(),
        )
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(entry.model_dump_json() + "\n")

    def save_snapshot(self, doc: DocumentNode) -> Path:
        """Saves a full state snapshot of the Canonical Document Model."""
        snap_path = self.snapshots_dir / f"snapshot_rev_{doc.version:06d}.json"
        with open(snap_path, "w", encoding="utf-8") as f:
            f.write(doc.model_dump_json(indent=2))
        return snap_path

    def load_latest_snapshot(self) -> Optional[DocumentNode]:
        """Loads the highest revision snapshot available."""
        snapshots = sorted(list(self.snapshots_dir.glob("snapshot_rev_*.json")))
        if not snapshots:
            return None
        latest = snapshots[-1]
        with open(latest, "r", encoding="utf-8") as f:
            data = json.load(f)
        return DocumentNode(**data)

    def has_recoverable_session(self) -> bool:
        """Returns True if unsaved recovery snapshots exist."""
        return len(list(self.snapshots_dir.glob("snapshot_rev_*.json"))) > 0

    def cleanup(self) -> None:
        """Cleans up local scratch persistence upon graceful exit/export."""
        if self.workspace_dir.exists():
            import shutil
            shutil.rmtree(self.workspace_dir, ignore_errors=True)
