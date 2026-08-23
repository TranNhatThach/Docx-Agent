"""
Semantic Diff Engine: Computes structural, textual, and formatting diffs between DOCX revisions.
"""

from pathlib import Path
from typing import Dict, Any, List, Optional, Union
import docx
from docx.document import Document as DocxDocument
from pydantic import BaseModel, Field
from docx_agent.core.document import DocumentModel
from docx_agent.utils.unicode import normalize_unicode


class ElementDiff(BaseModel):
    id: str
    type: str = "paragraph"
    status: str  # "modified", "added", "removed", "unchanged"
    before: Optional[str] = None
    after: Optional[str] = None
    format_changes: Dict[str, Any] = Field(default_factory=dict)


class DocumentDiffReport(BaseModel):
    file_before: Optional[str] = None
    file_after: Optional[str] = None
    identical: bool
    summary: Dict[str, int] = Field(default_factory=dict)
    changed: List[ElementDiff] = Field(default_factory=list)
    structural_changes: List[str] = Field(default_factory=list)


class DiffEngine:
    """
    Compares two DOCX documents and produces structured semantic diffs.
    """

    @staticmethod
    def compare(
        before_doc_or_path: Union[str, docx.Document],
        after_doc_or_path: Union[str, docx.Document],
    ) -> DocumentDiffReport:
        model_a = DocumentModel(before_doc_or_path)
        model_b = DocumentModel(after_doc_or_path)

        paras_a = model_a.get_paragraphs(include_runs=True)
        paras_b = model_b.get_paragraphs(include_runs=True)

        changed_list: List[ElementDiff] = []
        structural_changes: List[str] = []

        # Compare paragraph counts
        if len(paras_a) != len(paras_b):
            structural_changes.append(
                f"Paragraph count changed from {len(paras_a)} to {len(paras_b)}"
            )

        # Compare tables
        tables_a = len(model_a.doc.tables)
        tables_b = len(model_b.doc.tables)
        if tables_a != tables_b:
            structural_changes.append(f"Table count changed from {tables_a} to {tables_b}")

        # Compare sections
        secs_a = len(model_a.doc.sections)
        secs_b = len(model_b.doc.sections)
        if secs_a != secs_b:
            structural_changes.append(f"Section count changed from {secs_a} to {secs_b}")

        # Compare paragraphs side-by-side
        max_len = max(len(paras_a), len(paras_b))
        added_cnt = 0
        removed_cnt = 0
        modified_cnt = 0

        for i in range(max_len):
            pa = paras_a[i] if i < len(paras_a) else None
            pb = paras_b[i] if i < len(paras_b) else None

            if pa is not None and pb is not None:
                text_changed = pa.text != pb.text
                fmt_changes: Dict[str, Any] = {}

                if pa.style != pb.style:
                    fmt_changes["style"] = {"before": pa.style, "after": pb.style}
                if pa.alignment != pb.alignment:
                    fmt_changes["alignment"] = {"before": pa.alignment, "after": pb.alignment}

                if text_changed or fmt_changes:
                    modified_cnt += 1
                    changed_list.append(
                        ElementDiff(
                            id=pb.id,
                            type="paragraph",
                            status="modified",
                            before=pa.text,
                            after=pb.text,
                            format_changes=fmt_changes,
                        )
                    )
            elif pa is None and pb is not None:
                added_cnt += 1
                changed_list.append(
                    ElementDiff(
                        id=pb.id,
                        type="paragraph",
                        status="added",
                        before=None,
                        after=pb.text,
                    )
                )
            elif pa is not None and pb is None:
                removed_cnt += 1
                changed_list.append(
                    ElementDiff(
                        id=pa.id,
                        type="paragraph",
                        status="removed",
                        before=pa.text,
                        after=None,
                    )
                )

        is_identical = (
            len(changed_list) == 0 and len(structural_changes) == 0
        )

        return DocumentDiffReport(
            file_before=model_a.file_path,
            file_after=model_b.file_path,
            identical=is_identical,
            summary={
                "modified": modified_cnt,
                "added": added_cnt,
                "removed": removed_cnt,
                "structural": len(structural_changes),
            },
            changed=changed_list,
            structural_changes=structural_changes,
        )
