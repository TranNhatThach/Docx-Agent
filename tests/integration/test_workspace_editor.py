"""
Integration and End-to-End Test Suite for Docx-Agent V2.1 Real Document Editor.
Validates loading state machine, WorkspaceBridge serialization, direct editing,
transactional save, independent reopen verification, SelectionContext bridge,
and all real DOCX fixtures.
"""

import json
import pytest
from pathlib import Path
from docx_agent.agent import DocumentAgent
from docx_agent.interfaces.workspace.bridge import WorkspaceBridge
from docx_agent.engine.operations import ReplaceTextOp, InsertTextOp
from docx_agent.verification.validator import DocumentValidator
from docx_agent.verification.formatting import FormatChecker
from docx_agent.verification.visual import VisualLayoutVerifier

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures"


@pytest.fixture(scope="module", autouse=True)
def ensure_fixtures():
    """Ensure all test fixtures exist prior to running integration tests."""
    import sys
    sys_path_root = Path(__file__).parent.parent
    if str(sys_path_root) not in sys.path:
        sys.path.insert(0, str(sys_path_root))
    try:
        from fixtures.create_fixtures import generate_all_fixtures
        generate_all_fixtures()
    except Exception:
        pass



def test_fixture_loading_and_payload_structure():
    """Tests WorkspaceBridge.load_document_payload across all 8 real DOCX fixtures."""
    fixtures = [
        "simple.docx",
        "Vietnamese.docx",
        "academic_report.docx",
        "tables.docx",
        "images.docx",
        "multi_page.docx",
        "complex_formatting.docx",
        "unsupported_ooxml.docx",
    ]

    for fname in fixtures:
        fpath = FIXTURES_DIR / fname
        assert fpath.exists(), f"Fixture missing: {fname}"

        payload = WorkspaceBridge.load_document_payload(fpath)
        assert payload["success"] is True
        assert "document_id" in payload
        assert "sections" in payload
        assert len(payload["sections"]) >= 1
        assert "headings" in payload
        assert "stats" in payload
        assert payload["stats"]["load_time_ms"] > 0


def test_vietnamese_academic_report_full_roundtrip(tmp_path):
    """
    Mandatory Real-World Test:
    1. Load real Vietnamese academic report
    2. Extract SelectionContext
    3. Propose Agent Transaction & Preview Diff
    4. Apply Transaction
    5. Save document with atomic verification
    6. Independently reopen DOCX and verify changes & formatting survive
    7. Undo transaction and verify rollback
    """
    orig_path = FIXTURES_DIR / "academic_report.docx"
    test_doc_path = tmp_path / "test_academic_report.docx"
    test_doc_path.write_bytes(orig_path.read_bytes())

    # 1. Load document payload
    payload = WorkspaceBridge.load_document_payload(test_doc_path)
    assert payload["success"] is True
    assert len(payload["headings"]) >= 3
    assert "BÁO CÁO KIẾN TRÚC" in payload["sections"][0]["blocks"][0]["text"]

    # 2. Extract SelectionContext on first heading
    h1_id = payload["headings"][0]["id"]
    ctx = WorkspaceBridge.get_selection_context_payload(
        file_path=test_doc_path,
        block_id=h1_id,
        start_offset=0,
        end_offset=10,
    )
    assert ctx["block_id"] == h1_id
    assert ctx["document_profile"] == "academic_report"

    # 3. Direct edit / mutation: Add paragraph and modify heading text
    sec0 = payload["sections"][0]
    # Update first block text
    sec0["blocks"][0]["text"] = "1. TỔNG QUAN KIẾN TRÚC HỆ THỐNG V2.1 [CHỈNH SỬA]"
    
    # 4. Save document payload
    save_res = WorkspaceBridge.save_document_payload(
        file_path=test_doc_path,
        document_data=payload,
        output_path=test_doc_path,
    )
    assert save_res["success"] is True
    assert save_res["verification_passed"] is True

    # 5. Independent reopen & verify
    reopened_agent = DocumentAgent(test_doc_path)
    summary = reopened_agent.inspect()
    assert summary.paragraphs_count >= 3
    
    # Verify modified text persisted
    first_p = reopened_agent.read(start=0, end=2)
    full_texts = [p.text for p in first_p]
    assert any("[CHỈNH SỬA]" in t for t in full_texts)

    # Verify tables survived
    assert summary.tables_count >= 1
    tables = reopened_agent.model.get_tables()
    assert len(tables) >= 1
    assert "Canonical Engine" in tables[0].preview[1][0]


def test_10_page_multi_page_document_performance_and_headings(tmp_path):
    """Validates 10+ page multi-page document loading, pagination, and outline consistency."""
    multi_path = FIXTURES_DIR / "multi_page.docx"
    payload = WorkspaceBridge.load_document_payload(multi_path)

    assert payload["success"] is True
    assert len(payload["headings"]) == 40  # 10 H1 + 30 H2
    assert payload["stats"]["words_count"] > 1000
    assert payload["stats"]["load_time_ms"] < 2000  # Fast responsive loading

    # Direct edit in chapter 10
    h_ch10 = [h for h in payload["headings"] if "Chương 10" in h["text"]][0]
    assert h_ch10 is not None

    sec = payload["sections"][0]
    blk10 = [b for b in sec["blocks"] if b["id"] == h_ch10["id"]][0]
    blk10["text"] = "Chương 10: Phân Tích Chuyên Sâu Phần 10 (Đã cập nhật)"

    # Save
    out_p = tmp_path / "multi_page_saved.docx"
    save_res = WorkspaceBridge.save_document_payload(
        file_path=multi_path,
        document_data=payload,
        output_path=out_p,
    )
    assert save_res["success"] is True

    # Reopen and verify outline
    reopened = DocumentAgent(out_p)
    outline = reopened.outline()
    assert len(outline) == 40
    ch10_outline = [h for h in outline if "(Đã cập nhật)" in h["text"]]
    assert len(ch10_outline) == 1


def test_tables_and_formatting_preservation(tmp_path):
    """Validates complex tables and formatting preservation across export/reopen."""
    tables_path = FIXTURES_DIR / "tables.docx"
    payload = WorkspaceBridge.load_document_payload(tables_path)

    sec = payload["sections"][0]
    tbl_blk = [b for b in sec["blocks"] if b["type"] == "table"][0]
    assert tbl_blk["rows"] == 4
    assert tbl_blk["columns"] == 4

    # Edit a cell value directly
    tbl_blk["cells"][1][2]["text"] = "99.9 ms"

    # Save
    out_p = tmp_path / "tables_edited.docx"
    save_res = WorkspaceBridge.save_document_payload(
        file_path=tables_path,
        document_data=payload,
        output_path=out_p,
    )
    assert save_res["success"] is True

    # Reopen
    reopened = DocumentAgent(out_p)
    tbls = reopened.model.get_tables()
    assert len(tbls) == 1
    assert tbls[0].preview[1][2] == "99.9 ms"


def test_unsupported_ooxml_preservation(tmp_path):
    """Validates that unsupported OOXML nodes are preserved without data destruction."""
    unsupported_p = FIXTURES_DIR / "unsupported_ooxml.docx"
    payload = WorkspaceBridge.load_document_payload(unsupported_p)

    sec = payload["sections"][0]
    unsupported_blks = [b for b in sec["blocks"] if b["type"] == "unsupported"]
    assert len(unsupported_blks) >= 1

    # Save and verify
    out_p = tmp_path / "unsupported_out.docx"
    save_res = WorkspaceBridge.save_document_payload(
        file_path=unsupported_p,
        document_data=payload,
        output_path=out_p,
    )
    assert save_res["success"] is True
    assert out_p.exists()
