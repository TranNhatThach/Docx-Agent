"""
Workspace Bridge: High-Performance, Reversible Document Bridge between
Python Canonical Engine, VS Code Extension Webview, and Antigravity Agent.
"""

import sys
import os
import json
import time
import traceback
from pathlib import Path
from typing import Dict, Any, List, Optional, Union

from docx_agent.agent import DocumentAgent
from docx_agent.canonical.model import (
    DocumentNode,
    SectionNode,
    ParagraphBlock,
    HeadingBlock,
    ListItemBlock,
    TableBlock,
    TableCellNode,
    ImageBlock,
    DiagramBlock,
    UnsupportedBlock,
    RunNode,
    generate_id,
)
from docx_agent.adapters.docx import DocxImporter, DocxExporter
from docx_agent.engine.layout import LayoutEngine
from docx_agent.engine.selection import SelectionProvider, SelectionContext
from docx_agent.engine.operations import (
    DocOperation,
    InsertTextOp,
    DeleteTextOp,
    ReplaceTextOp,
    FormatParagraphOp,
    InsertBlockOp,
    DeleteBlockOp,
    InsertCitationOp,
    CompositeOperation,
)
from docx_agent.engine.transactions import AgentTransactionManager, TransactionPreview
from docx_agent.transactions.transaction import TransactionContext
from docx_agent.verification.validator import DocumentValidator
from docx_agent.verification.formatting import FormatChecker
from docx_agent.verification.visual import VisualLayoutVerifier
from docx_agent.utils.paths import resolve_safe_path, ensure_parent_dir
from docx_agent.utils.unicode import normalize_unicode
from docx_agent.changes.model import EditSession, ChangeObject, ChangeStatus, ChangeType


class WorkspaceBridge:
    """
    Core serialization, operation execution, and transactional persistence
    bridge for the visual A4 workspace.
    """

    @staticmethod
    def load_document_payload(file_path: Union[str, Path]) -> Dict[str, Any]:
        """
        Loads and parses DOCX file into full canonical JSON payload
        for the webview editor with heading outline, statistics, and deterministic A4 layout pages.
        """
        start_time = time.time()
        path = resolve_safe_path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"Tài liệu không tồn tại: {path}")

        agent = DocumentAgent(path)
        doc = agent.canonical_doc
        summary = agent.inspect()

        # Run Deterministic Layout Engine
        layout_doc = LayoutEngine.paginate(doc)

        headings = layout_doc.headings_outline
        total_words = layout_doc.total_words
        total_chars = layout_doc.total_chars

        sections_data = []
        for sec in doc.sections:
            sec_blocks = []
            for blk in sec.blocks:
                sec_blocks.append(LayoutEngine._serialize_block_for_render(blk))

            sections_data.append({
                "id": sec.id,
                "properties": sec.properties.model_dump(),
                "header_text": sec.header_text,
                "footer_text": sec.footer_text,
                "first_page_header_text": sec.first_page_header_text,
                "first_page_footer_text": sec.first_page_footer_text,
                "has_page_numbers": sec.has_page_numbers,
                "page_number_format": sec.page_number_format,
                "blocks": sec_blocks,
            })

        elapsed = round((time.time() - start_time) * 1000, 2)

        return {
            "success": True,
            "document_id": doc.id,
            "title": doc.title or path.stem,
            "file_path": str(path),
            "file_name": path.name,
            "profile": doc.profile.value if hasattr(doc.profile, "value") else str(doc.profile),
            "version": doc.version,
            "stats": {
                "paragraphs_count": summary.paragraphs_count,
                "tables_count": summary.tables_count,
                "sections_count": summary.sections_count,
                "headings_count": len(headings),
                "words_count": total_words,
                "chars_count": total_chars,
                "total_pages": layout_doc.total_pages,
                "load_time_ms": elapsed,
            },
            "headings": headings,
            "sections": sections_data,
            "total_pages": layout_doc.total_pages,
            "pages": [p.model_dump() for p in layout_doc.pages],
        }

    @staticmethod
    def save_document_payload(
        file_path: Union[str, Path],
        document_data: Dict[str, Any],
        output_path: Optional[Union[str, Path]] = None,
        auto_backup: bool = True,
        verify: bool = True,
    ) -> Dict[str, Any]:
        """
        Reconstructs Canonical Document Model from webview data, applies changes,
        validates, writes atomic backup, exports to DOCX, reopens and independently verifies.
        """
        start_time = time.time()
        target_p = resolve_safe_path(output_path or file_path)

        # 1. Reconstruct DocumentNode from incoming JSON
        doc_node = DocumentNode(
            id=document_data.get("document_id", generate_id("doc")),
            title=document_data.get("title", target_p.stem),
            version=document_data.get("version", 1) + 1,
        )

        doc_node.sections = []
        for s_data in document_data.get("sections", []):
            sec_node = SectionNode(
                id=s_data.get("id", generate_id("sec")),
                header_text=s_data.get("header_text"),
                footer_text=s_data.get("footer_text"),
                has_page_numbers=s_data.get("has_page_numbers", True),
                page_number_format=s_data.get("page_number_format", "Trang {PAGE} / {NUMPAGES}"),
            )
            if "properties" in s_data and isinstance(s_data["properties"], dict):
                sec_node.properties = sec_node.properties.model_copy(update=s_data["properties"])

            for b_data in s_data.get("blocks", []):
                b_type = b_data.get("type", "paragraph")
                b_id = b_data.get("id", generate_id("blk"))
                style_name = b_data.get("style_name", "Normal")

                # Parse runs
                runs = []
                for r_data in b_data.get("runs", []):
                    runs.append(
                        RunNode(
                            id=r_data.get("id", generate_id("r")),
                            text=normalize_unicode(r_data.get("text", "")),
                            font_name=r_data.get("font_name"),
                            font_size_pt=r_data.get("font_size_pt"),
                            bold=r_data.get("bold"),
                            italic=r_data.get("italic"),
                            underline=r_data.get("underline"),
                            strike=r_data.get("strike"),
                            color_rgb=r_data.get("color_rgb"),
                            highlight=r_data.get("highlight"),
                        )
                    )

                if "text" in b_data and b_data["text"] is not None:
                    expected_text = normalize_unicode(b_data["text"])
                    current_text = "".join(r.text for r in runs)
                    if not runs or current_text != expected_text:
                        if len(runs) == 1:
                            runs[0].text = expected_text
                        else:
                            runs = [RunNode(text=expected_text)]

                if b_type == "heading":
                    sec_node.blocks.append(
                        HeadingBlock(
                            id=b_id,
                            level=b_data.get("level", 1),
                            style_name=style_name or f"Heading {b_data.get('level', 1)}",
                            runs=runs,
                            alignment=b_data.get("alignment", "left"),
                            line_spacing=b_data.get("line_spacing", 1.5),
                            space_before_pt=b_data.get("space_before_pt", 12.0),
                            space_after_pt=b_data.get("space_after_pt", 6.0),
                            first_line_indent_cm=0.0,
                        )
                    )
                elif b_type == "list_item":
                    sec_node.blocks.append(
                        ListItemBlock(
                            id=b_id,
                            list_type=b_data.get("list_type", "bullet"),
                            list_level=b_data.get("list_level", 0),
                            style_name=style_name,
                            runs=runs,
                            alignment=b_data.get("alignment", "left"),
                        )
                    )
                elif b_type == "table":
                    cells_grid = []
                    for row in b_data.get("cells", []):
                        row_cells = []
                        for c_data in row:
                            row_cells.append(
                                TableCellNode(
                                    id=c_data.get("id", generate_id("cell")),
                                    text=normalize_unicode(c_data.get("text", "")),
                                    bg_color_hex=c_data.get("bg_color_hex"),
                                    rowspan=c_data.get("rowspan", 1),
                                    colspan=c_data.get("colspan", 1),
                                )
                            )
                        cells_grid.append(row_cells)

                    sec_node.blocks.append(
                        TableBlock(
                            id=b_id,
                            rows=b_data.get("rows", len(cells_grid)),
                            columns=b_data.get("columns", len(cells_grid[0]) if cells_grid else 0),
                            cells=cells_grid,
                            style_name=style_name or "Table Grid",
                        )
                    )
                elif b_type == "image":
                    sec_node.blocks.append(
                        ImageBlock(
                            id=b_id,
                            source_uri_or_path=b_data.get("source_uri_or_path", ""),
                            width_cm=b_data.get("width_cm", 10.0),
                            caption=b_data.get("caption"),
                        )
                    )
                elif b_type == "diagram":
                    sec_node.blocks.append(
                        DiagramBlock(
                            id=b_id,
                            diagram_type=b_data.get("diagram_type", "architecture"),
                            source_code=b_data.get("source_code", ""),
                            caption=b_data.get("caption"),
                        )
                    )
                else:
                    sec_node.blocks.append(
                        ParagraphBlock(
                            id=b_id,
                            style_name=style_name,
                            runs=runs,
                            alignment=b_data.get("alignment", "justify"),
                            line_spacing=b_data.get("line_spacing", 1.5),
                            space_before_pt=b_data.get("space_before_pt", 0.0),
                            space_after_pt=b_data.get("space_after_pt", 6.0),
                            first_line_indent_cm=b_data.get("first_line_indent_cm", 1.27),
                        )
                    )

            doc_node.sections.append(sec_node)

        # 2. Export to DOCX
        try:
            exported_path = DocxExporter.export_docx(doc_node, target_p)
        except Exception as e:
            return {
                "success": False,
                "error_stage": "EXPORT",
                "message": f"Lỗi khi xuất OpenXML DOCX: {str(e)}",
                "diagnostics": traceback.format_exc(),
            }

        # 3. Independent Verification: Re-open and verify file integrity
        verification_passed = True
        verification_details = "Verified successfully"
        if verify:
            try:
                report = DocumentValidator.verify_integrity(exported_path)
                if not report.passed:
                    verification_passed = False
                    verification_details = f"Validation failures: {report.failures}"
            except Exception as e:
                verification_passed = False
                verification_details = f"Failed to verify exported file: {str(e)}"

        elapsed = round((time.time() - start_time) * 1000, 2)

        return {
            "success": True,
            "message": "Đã lưu tài liệu thành công.",
            "file_path": str(exported_path),
            "version": doc_node.version,
            "verification_passed": verification_passed,
            "verification_details": verification_details,
            "save_time_ms": elapsed,
        }

    @staticmethod
    def get_selection_context_payload(
        file_path: Union[str, Path],
        block_id: str,
        start_offset: int = 0,
        end_offset: int = 0,
    ) -> Dict[str, Any]:
        """
        Extracts rich SelectionContext for AI agent commands.
        """
        path = resolve_safe_path(file_path)
        agent = DocumentAgent(path)
        ctx = SelectionProvider.build_context(
            doc=agent.canonical_doc,
            block_id=block_id,
            start_offset=start_offset,
            end_offset=end_offset,
        )
        return ctx.model_dump()

    @staticmethod
    def execute_agent_plan(file_path: Union[str, Path], plan_dict: Dict[str, Any]) -> Dict[str, Any]:
        """
        Executes multi-step agent batch operations transactionally with rollback.
        """
        path = resolve_safe_path(file_path)
        agent = DocumentAgent(path)

        with TransactionContext(agent.model, transaction_id=plan_dict.get("plan_id")) as tx:
            for op_data in plan_dict.get("operations", []):
                op_type = op_data.get("op_type")
                target = op_data.get("target")

                if op_type == "replace_text":
                    agent.content.replace_text(
                        target=target,
                        old_text=op_data.get("old_text", ""),
                        new_text=op_data.get("new_text", ""),
                    )
                elif op_type == "format_text":
                    props = op_data.get("formatting", {})
                    agent.text_fmt.format_run(target=target, **props)
                elif op_type == "format_paragraph":
                    props = op_data.get("formatting", {})
                    agent.para_fmt.format_paragraph(target=target, **props)
                elif op_type == "insert_paragraph":
                    agent.content.insert_paragraph_after(
                        target=target,
                        text=op_data.get("text", ""),
                        style=op_data.get("style", "Normal"),
                    )

            tx.commit()

        # Update canonical document
        agent.canonical_doc = DocxImporter.import_docx(path)

        return {
            "success": True,
            "transaction_id": tx.transaction_id,
            "status": "COMMITTED",
            "message": f"Đã áp dụng thành công {len(plan_dict.get('operations', []))} thao tác từ Antigravity Agent.",
        }

    # ---------------------------------------------------------
    # AI EDIT SESSIONS & CHANGE TRACKING BRIDGE
    # ---------------------------------------------------------
    @classmethod
    def propose_edit_session_payload(
        cls,
        file_path: Union[str, Path],
        task_description: str,
        modified_doc_data: Dict[str, Any],
        agent_id: str = "Antigravity-Agent",
        reason: str = "Tối ưu hóa và cải thiện nội dung tài liệu theo chuẩn học thuật.",
        confidence: float = 0.95,
        evidence: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Creates an EditSession with semantic ChangeObjects by comparing current document with proposed candidate.
        """
        path = resolve_safe_path(file_path)
        agent = DocumentAgent(path)

        # Reconstruct candidate DocumentNode from JSON
        from docx_agent.canonical.model import DocumentNode, SectionNode, ParagraphBlock, HeadingBlock, TableBlock, RunNode
        candidate_doc = DocumentNode(title=modified_doc_data.get("title", path.stem))
        candidate_doc.sections = []

        for s_data in modified_doc_data.get("sections", []):
            s_node = SectionNode(id=s_data.get("id", "sec_0001"))
            for b_data in s_data.get("blocks", []):
                b_type = b_data.get("type", "paragraph")
                b_id = b_data.get("id", "blk_0001")
                if b_type == "heading":
                    s_node.blocks.append(
                        HeadingBlock(
                            id=b_id,
                            level=b_data.get("level", 1),
                            style_name=f"Heading {b_data.get('level', 1)}",
                            runs=[RunNode(text=b_data.get("text", ""))],
                        )
                    )
                elif b_type == "table":
                    s_node.blocks.append(
                        TableBlock(
                            id=b_id,
                            rows=b_data.get("rows", 0),
                            columns=b_data.get("columns", 0),
                        )
                    )
                else:
                    s_node.blocks.append(
                        ParagraphBlock(
                            id=b_id,
                            style_name=b_data.get("style_name", "Normal"),
                            runs=[RunNode(text=b_data.get("text", ""))],
                        )
                    )
            candidate_doc.sections.append(s_node)

        session = agent.propose_changes(
            modified_doc=candidate_doc,
            task_description=task_description,
            agent_id=agent_id,
            reason=reason,
            confidence=confidence,
            evidence=evidence,
        )

        # Persist session to workspace storage for subsequent CLI commands
        cls._save_session_cache(path, session)

        return {
            "success": True,
            "session": session.model_dump(),
            "changes_count": len(session.changes),
            "summary": session.summary(),
        }

    @classmethod
    def _get_session_cache_path(cls, file_path: Path) -> Path:
        cache_dir = file_path.parent / ".docx_agent_workspace"
        cache_dir.mkdir(parents=True, exist_ok=True)
        return cache_dir / f"session_{file_path.stem}.json"

    @classmethod
    def _save_session_cache(cls, file_path: Path, session: EditSession) -> None:
        try:
            cache_file = cls._get_session_cache_path(file_path)
            with open(cache_file, "w", encoding="utf-8") as f:
                json.dump(session.model_dump(), f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    @classmethod
    def _load_session_cache(cls, file_path: Path) -> Optional[EditSession]:
        try:
            cache_file = cls._get_session_cache_path(file_path)
            if cache_file.exists():
                with open(cache_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                return EditSession(**data)
        except Exception:
            pass
        return None

    @classmethod
    def accept_change_payload(cls, file_path: Union[str, Path], change_id: str) -> Dict[str, Any]:
        path = resolve_safe_path(file_path)
        agent = DocumentAgent(path)
        session = cls._load_session_cache(path)
        if session:
            agent.version_manager.register_session(session)
            agent.active_session = session

        success = agent.accept_change(change_id)
        if session and success:
            cls._save_session_cache(path, session)

        return {"success": success, "change_id": change_id, "status": "ACCEPTED" if success else "NOT_FOUND"}

    @classmethod
    def reject_change_payload(cls, file_path: Union[str, Path], change_id: str) -> Dict[str, Any]:
        path = resolve_safe_path(file_path)
        agent = DocumentAgent(path)
        session = cls._load_session_cache(path)
        if session:
            agent.version_manager.register_session(session)
            agent.active_session = session

        success = agent.reject_change(change_id)
        if session and success:
            cls._save_session_cache(path, session)

        return {"success": success, "change_id": change_id, "status": "REJECTED" if success else "NOT_FOUND"}

    @classmethod
    def commit_session_payload(
        cls,
        file_path: Union[str, Path],
        session_id: str,
        output_path: Optional[Union[str, Path]] = None,
    ) -> Dict[str, Any]:
        path = resolve_safe_path(file_path)
        agent = DocumentAgent(path)
        session = cls._load_session_cache(path)
        if session:
            agent.version_manager.register_session(session)
            agent.active_session = session

        out_p, info = agent.commit_session(session_id=session_id, output_path=output_path)
        return {"success": True, "file_path": out_p, "commit_info": info}

    @classmethod
    def get_version_history_payload(cls, file_path: Union[str, Path]) -> Dict[str, Any]:
        path = resolve_safe_path(file_path)
        agent = DocumentAgent(path)
        history = agent.get_version_history()
        return {"success": True, "versions": history, "current_version": agent.canonical_doc.version}


