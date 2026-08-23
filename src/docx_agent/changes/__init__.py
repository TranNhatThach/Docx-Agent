"""
DOCX Agent Semantic Change Tracking & Versioning Subsystem.
"""

from docx_agent.changes.model import (
    ChangeObject,
    ChangeType,
    ChangeStatus,
    ChangeLocation,
    EditSession,
    generate_change_id,
    generate_session_id,
)
from docx_agent.changes.diff import SemanticDiffEngine
from docx_agent.changes.versioning import VersionSnapshot, VersionManager
from docx_agent.changes.revisions import NativeWordRevisionExporter

__all__ = [
    "ChangeObject",
    "ChangeType",
    "ChangeStatus",
    "ChangeLocation",
    "EditSession",
    "generate_change_id",
    "generate_session_id",
    "SemanticDiffEngine",
    "VersionSnapshot",
    "VersionManager",
    "NativeWordRevisionExporter",
]
