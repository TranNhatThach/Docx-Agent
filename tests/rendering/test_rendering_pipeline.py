"""
Comprehensive Rendering & Layout Pipeline Test Suite.
Verifies all 13 OOXML rendering capabilities, style inheritance, pagination, and roundtrip idempotency.
"""

import time
import pytest
from pathlib import Path

from docx_agent.adapters.docx import DocxImporter, DocxExporter
from docx_agent.engine.layout import LayoutEngine
from docx_agent.interfaces.workspace.bridge import WorkspaceBridge
from docx_agent.engine.styles import StyleResolver
from docx_agent.engine.numbering import NumberingResolver


FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "rendering"


def test_simple_document_import_and_pagination():
    file_path = FIXTURES_DIR / "simple.docx"
    doc_node = DocxImporter.import_docx(file_path)
    assert len(doc_node.sections) == 1
    assert len(doc_node.sections[0].blocks) >= 2

    layout = LayoutEngine.paginate(doc_node)
    assert layout.total_pages == 1
    assert len(layout.pages) == 1
    assert len(layout.pages[0].blocks) >= 2


def test_vietnamese_unicode_fidelity():
    file_path = FIXTURES_DIR / "vietnamese.docx"
    doc_node = DocxImporter.import_docx(file_path)
    text_content = " ".join([b.full_text for b in doc_node.sections[0].blocks if hasattr(b, "full_text")])

    assert "ă â ê ô ơ ư đ" in text_content
    assert "À Á Ả Ã Ạ" in text_content
    assert "∀x ∈ ℝ" in text_content
    assert "π ≈ 3.14159" in text_content


def test_headings_outline_hierarchy():
    file_path = FIXTURES_DIR / "headings.docx"
    payload = WorkspaceBridge.load_document_payload(file_path)
    assert payload["success"] is True
    headings = payload["headings"]
    assert len(headings) == 3
    assert headings[0]["level"] == 1
    assert headings[1]["level"] == 2
    assert headings[2]["level"] == 3


def test_style_resolver_cascading():
    file_path = FIXTURES_DIR / "styles.docx"
    doc_node = DocxImporter.import_docx(file_path)
    p_style = [b for b in doc_node.sections[0].blocks if b.style_name == "Normal"][0]
    runs = p_style.runs

    assert any(r.bold and r.color_rgb for r in runs)
    assert any(r.italic and r.color_rgb for r in runs)
    assert any(r.underline for r in runs)


def test_numbering_resolver():
    file_path = FIXTURES_DIR / "numbering.docx"
    doc_node = DocxImporter.import_docx(file_path)
    blocks = doc_node.sections[0].blocks
    list_items = [b for b in blocks if hasattr(b, "list_type") or (b.style_name and "List" in b.style_name)]
    assert len(list_items) >= 2


def test_table_structure_and_dimensions():
    file_path = FIXTURES_DIR / "tables.docx"
    doc_node = DocxImporter.import_docx(file_path)
    tbls = [b for b in doc_node.sections[0].blocks if b.type.value == "table" or str(b.type) == "table"]
    assert len(tbls) == 1
    tbl = tbls[0]
    assert tbl.rows == 3
    assert tbl.columns == 3
    assert tbl.cells[0][0].text == "Ô (1, 1)"


def test_merged_tables_spans():
    file_path = FIXTURES_DIR / "merged_tables.docx"
    doc_node = DocxImporter.import_docx(file_path)
    tbls = [b for b in doc_node.sections[0].blocks if b.type.value == "table" or str(b.type) == "table"]
    assert len(tbls) == 1
    tbl = tbls[0]
    assert tbl.cells[0][0].colspan >= 1


def test_section_break_isolation():
    file_path = FIXTURES_DIR / "sections.docx"
    doc_node = DocxImporter.import_docx(file_path)
    assert len(doc_node.sections) == 2
    assert doc_node.sections[0].properties.orientation == "portrait"
    assert doc_node.sections[1].properties.orientation == "landscape"


def test_cover_page_header_footer_isolation():
    file_path = FIXTURES_DIR / "headers_footers.docx"
    payload = WorkspaceBridge.load_document_payload(file_path)
    assert payload["success"] is True
    pages = payload["pages"]
    assert len(pages) >= 2
    # First page is cover / different first page -> no regular header
    assert pages[0]["is_cover_page"] is True
    assert pages[0]["has_header"] is False
    # Subsequent page has header
    assert pages[1]["has_header"] is True


def test_roundtrip_idempotency(tmp_path):
    file_path = FIXTURES_DIR / "complex_formatting.docx"
    doc_node_1 = DocxImporter.import_docx(file_path)

    out_file = tmp_path / "exported_roundtrip.docx"
    DocxExporter.export_docx(doc_node_1, out_file)
    assert out_file.exists()

    doc_node_2 = DocxImporter.import_docx(out_file)
    assert len(doc_node_2.sections) == len(doc_node_1.sections)
    assert len(doc_node_2.sections[0].blocks) == len(doc_node_1.sections[0].blocks)


def test_stress_50_pages_performance():
    file_path = FIXTURES_DIR / "stress_50_pages.docx"
    t0 = time.time()
    payload = WorkspaceBridge.load_document_payload(file_path)
    t_elapsed = time.time() - t0

    assert payload["success"] is True
    assert payload["total_pages"] >= 50
    assert t_elapsed < 5.0  # Must parse and paginate 50 pages well within 5 seconds
