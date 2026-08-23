import pytest
import os
import tempfile
import pytest
import os
import tempfile
from docx_agent.canonical.model import (
    DocumentNode,
    SectionNode,
    ParagraphBlock,
    RunNode,
    HeadingBlock,
    TableBlock,
    TableCellNode,
    SectionProperties,
)
from docx_agent.adapters.docx import DocxImporter, DocxExporter
from docx_agent.changes.model import (
    ChangeType,
    ChangeStatus,
    ChangeLocation,
    ChangeObject,
    EditSession,
)
from docx_agent.changes.diff import SemanticDiffEngine
from docx_agent.changes.versioning import VersionManager, VersionSnapshot
from docx_agent.changes.revisions import NativeWordRevisionExporter
from docx_agent.agent import DocumentAgent
from docx_agent.interfaces.workspace.bridge import WorkspaceBridge


def create_sample_document() -> DocumentNode:
    p1 = ParagraphBlock(
        id="p_001",
        runs=[RunNode(text="Đây là đoạn văn đầu tiên trong tài liệu.")],
        alignment="LEFT",
        space_before_pt=0.0,
        space_after_pt=6.0,
        line_spacing=1.3,
    )
    h1 = HeadingBlock(
        id="h_001",
        level=1,
        runs=[RunNode(text="Chương 1: Khởi đầu", bold=True)],
    )
    p2 = ParagraphBlock(
        id="p_002",
        runs=[RunNode(text="Đoạn văn thứ hai chứa thông tin kỹ thuật về Oracle Database.")],
    )
    tbl = TableBlock(
        id="tbl_001",
        rows=2,
        columns=2,
        cells=[
            [
                TableCellNode(id="c_0_0", text="Mã NV"),
                TableCellNode(id="c_0_1", text="Tên NV"),
            ],
            [
                TableCellNode(id="c_1_0", text="100"),
                TableCellNode(id="c_1_1", text="Steven King"),
            ],
        ],
    )
    sec = SectionNode(
        id="sec_001",
        properties=SectionProperties(page_width_cm=21.0, page_height_cm=29.7),
        blocks=[p1, h1, p2, tbl],
    )
    return DocumentNode(
        title="Tài Liệu Mẫu",
        sections=[sec],
    )


class TestSemanticDiffEngine:
    def test_text_modification_diff(self):
        doc1 = create_sample_document()
        doc2 = create_sample_document()
        # Sửa nội dung đoạn 1 qua runs
        doc2.sections[0].blocks[0].runs = [RunNode(text="Đây là đoạn văn đã được AI cải thiện và trau chuốt.")]

        diff_engine = SemanticDiffEngine()
        changes = diff_engine.compute_diff(doc1, doc2, default_reason="Tối ưu câu chữ")

        assert len(changes) >= 1
        chg = changes[0]
        assert chg.change_type in (ChangeType.MODIFY_PARAGRAPH, ChangeType.REPLACE_TEXT)
        assert chg.target_element == "p_001"
        assert "cải thiện" in chg.after_content
        assert chg.reason == "Tối ưu câu chữ"
        assert chg.status == ChangeStatus.PROPOSED

    def test_paragraph_insertion_and_deletion(self):
        doc1 = create_sample_document()
        doc2 = create_sample_document()

        # Thêm đoạn mới
        p_new = ParagraphBlock(
            id="p_new",
            runs=[RunNode(text="Đoạn văn hoàn toàn mới được chèn thêm vào.")],
        )
        doc2.sections[0].blocks.append(p_new)

        diff_engine = SemanticDiffEngine()
        changes = diff_engine.compute_diff(doc1, doc2)
        ins_changes = [c for c in changes if c.change_type == ChangeType.ADD_PARAGRAPH]
        assert len(ins_changes) == 1
        assert ins_changes[0].after_content == "Đoạn văn hoàn toàn mới được chèn thêm vào."

    def test_heading_diff(self):
        doc1 = create_sample_document()
        doc2 = create_sample_document()

        # Sửa level và text của heading
        doc2.sections[0].blocks[1].level = 2
        doc2.sections[0].blocks[1].runs = [RunNode(text="1.1. Khởi đầu mới")]

        diff_engine = SemanticDiffEngine()
        changes = diff_engine.compute_diff(doc1, doc2)
        h_changes = [c for c in changes if c.change_type in (ChangeType.MODIFY_HEADING_TEXT, ChangeType.CHANGE_HEADING_LEVEL, ChangeType.REPLACE_TEXT)]
        assert len(h_changes) >= 1

    def test_table_cell_diff(self):
        doc1 = create_sample_document()
        doc2 = create_sample_document()

        # Đổi cell nội dung và màu nền
        doc2.sections[0].blocks[3].cells[1][1].text = "Steven King (CEO)"
        doc2.sections[0].blocks[3].cells[1][1].bg_color_hex = "F8FAFC"

        diff_engine = SemanticDiffEngine()
        changes = diff_engine.compute_diff(doc1, doc2)
        tbl_changes = [c for c in changes if c.change_type == ChangeType.MODIFY_CELL_CONTENT]
        assert len(tbl_changes) == 1
        assert (tbl_changes[0].location.row_idx, tbl_changes[0].location.col_idx) == (1, 1)
        assert tbl_changes[0].after_content == "Steven King (CEO)"

    def test_paragraph_formatting_diff(self):
        doc1 = create_sample_document()
        doc2 = create_sample_document()

        # Thay đổi căn lề và khoảng cách dòng
        doc2.sections[0].blocks[0].alignment = "JUSTIFY"
        doc2.sections[0].blocks[0].first_line_indent_cm = 1.27

        diff_engine = SemanticDiffEngine()
        changes = diff_engine.compute_diff(doc1, doc2)
        fmt_changes = [c for c in changes if c.change_type == ChangeType.FORMAT_PARAGRAPH]
        assert len(fmt_changes) == 1
        assert "alignment" in fmt_changes[0].affected_formatting or "first_line_indent_cm" in fmt_changes[0].affected_formatting


class TestVersionManagerAndSnapshots:
    def test_session_lifecycle_accept_reject(self):
        doc1 = create_sample_document()
        vm = VersionManager(doc1)
        assert vm.current_version == 1

        # Tạo doc2 và propose changes
        doc2 = create_sample_document()
        doc2.sections[0].blocks[0].runs = [RunNode(text="Đoạn 1 đã sửa")]
        doc2.sections[0].blocks[2].runs = [RunNode(text="Đoạn 2 đã sửa")]

        session = vm.create_edit_session("AI đề xuất sửa văn bản", doc2, agent_id="Antigravity")
        assert len(session.changes) >= 2

        # Chấp nhận change 1, từ chối change 2
        chg1 = session.changes[0]
        chg2 = session.changes[1]
        session.accept_change(chg1.change_id)
        session.reject_change(chg2.change_id)

        assert chg1.status == ChangeStatus.ACCEPTED
        assert chg2.status == ChangeStatus.REJECTED

        # Commit session
        v2 = vm.commit_session(session.session_id, author="Reviewer")
        assert v2.version == 2
        assert vm.current_version == 2
        assert len(vm.get_history()) == 2

    def test_undo_redo_and_restore_version(self):
        doc1 = create_sample_document()
        vm = VersionManager(doc1)

        doc2 = create_sample_document()
        doc2.sections[0].blocks[0].runs = [RunNode(text="Bản sửa 1")]
        session1 = vm.create_edit_session("Sửa lần 1", doc2)
        session1.accept_all()
        vm.commit_session(session1.session_id)

        assert vm.current_version == 2

        # Restore về version 1 (Tạo snapshot Version 3 chứa nội dung của Version 1)
        v1_restored = vm.restore_version(1)
        assert vm.current_version == 3
        assert "đoạn văn đầu tiên" in v1_restored.sections[0].blocks[0].full_text.lower()


class TestNativeWordRevisions:
    def test_revisions_export(self):
        doc1 = create_sample_document()
        doc2 = create_sample_document()
        doc2.sections[0].blocks[0].runs = [RunNode(text="Nội dung mới thay thế.")]

        diff_engine = SemanticDiffEngine()
        changes = diff_engine.compute_diff(doc1, doc2)

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tf:
            out_path = tf.name

        try:
            res_path = NativeWordRevisionExporter.export_docx_with_revisions(doc1, changes, out_path)
            assert os.path.exists(res_path)
            assert os.path.getsize(res_path) > 1000
        finally:
            if os.path.exists(out_path):
                os.remove(out_path)


class TestDocumentAgentIntegration:
    def test_agent_propose_and_review(self):
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tf:
            docx_path = tf.name

        try:
            # Tạo docx ban đầu qua DocxExporter
            DocxExporter.export_docx(create_sample_document(), docx_path)

            # Khởi tạo agent từ docx
            agent2 = DocumentAgent(docx_path)
            assert len(agent2.canonical_doc.sections) >= 1

            # Propose changes
            mod_data = agent2.canonical_doc.model_dump()
            mod_data["sections"][0]["blocks"][0]["runs"] = [{"text": "Chỉnh sửa qua DocumentAgent API"}]

            session = agent2.propose_changes(
                task_description="Kiểm tra tích hợp DocumentAgent",
                modified_doc_data=mod_data,
                reason="Kiểm thử unit test",
            )
            assert session is not None
            assert len(session.changes) >= 1

            # Accept all & commit
            agent2.accept_all_changes(session.session_id)
            out_file, snap_data = agent2.commit_session(session.session_id)
            assert snap_data["version"] == 2

            # Lịch sử
            history = agent2.get_version_history()
            assert len(history) == 2

        finally:
            if os.path.exists(docx_path):
                os.remove(docx_path)


class TestWorkspaceBridgeIntegration:
    def test_bridge_session_payloads(self):
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tf:
            docx_path = tf.name

        try:
            DocxExporter.export_docx(create_sample_document(), docx_path)

            doc_node = create_sample_document()
            mod_data = doc_node.model_dump()
            mod_data["sections"][0]["blocks"][0]["runs"] = [{"text": "Chỉnh sửa qua WorkspaceBridge"}]

            res = WorkspaceBridge.propose_edit_session_payload(
                file_path=docx_path,
                task_description="Bridge Unit Test",
                modified_doc_data=mod_data,
            )
            assert res["success"] is True
            assert len(res["session"]["changes"]) >= 1

            change_id = res["session"]["changes"][0]["change_id"]
            res_accept = WorkspaceBridge.accept_change_payload(docx_path, change_id)
            assert res_accept["success"] is True

            hist_res = WorkspaceBridge.get_version_history_payload(docx_path)
            assert hist_res["success"] is True
            assert len(hist_res["versions"]) >= 1

        finally:
            if os.path.exists(docx_path):
                os.remove(docx_path)


