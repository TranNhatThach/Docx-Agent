"""
Comprehensive V2 Scenarios Test Suite (Scenarios A through J).
Validates Canonical Document Model, Operation Engine, Selection Context, Research Assistant,
Diagrams, Transactions, Undo/Redo, Persistence, and Unsupported Element Preservation.
"""

import pytest
from pathlib import Path
from docx_agent.agent import DocumentAgent
from docx_agent.canonical.model import (
    DocumentNode,
    SectionNode,
    ParagraphBlock,
    HeadingBlock,
    TableBlock,
    TableCellNode,
    ImageBlock,
    DiagramBlock,
    UnsupportedBlock,
    RunNode,
    SourceMetadata,
)
from docx_agent.engine.operations import (
    InsertTextOp,
    DeleteTextOp,
    ReplaceTextOp,
    FormatParagraphOp,
    InsertBlockOp,
    DeleteBlockOp,
    UpdateTableCellOp,
    InsertCitationOp,
    CompositeOperation,
)
from docx_agent.engine.transactions import AgentTransactionManager
from docx_agent.engine.persistence import WorkspacePersistence
from docx_agent.engine.selection import SelectionProvider
from docx_agent.research.provider import ResearchAssistant
from docx_agent.media.diagrams import DiagramSynthesizer
from docx_agent.engine.clarification import ClarificationEngine, AgentConfidence
from docx_agent.adapters.docx import DocxImporter, DocxExporter
from docx_agent.verification.visual import VisualLayoutVerifier


def test_scenario_a_human_edit_save_reopen_verify(sample_docx, tmp_path):
    """SCENARIO A: Direct human edit through operations, export DOCX, independent reopen & verify."""
    agent = DocumentAgent(sample_docx)
    blk_id = agent.canonical_doc.sections[0].blocks[0].id

    # Human types text
    ins_op = InsertTextOp(block_id=blk_id, offset=0, text="[Chỉnh Sửa] ")
    agent.tx_manager.execute_operation(ins_op)

    # Export canonical to DOCX
    out_path = tmp_path / "canonical_out.docx"
    agent.export_canonical(out_path)
    assert out_path.exists()

    # Independent reopen
    reopened_agent = DocumentAgent(out_path)
    summary = reopened_agent.inspect()
    assert summary.paragraphs_count >= 1
    assert "[Chỉnh Sửa]" in reopened_agent.read(start=0, end=1)[0].text


def test_scenario_b_selection_agent_preview_apply_undo(sample_docx):
    """SCENARIO B: Highlight text, build selection context, propose agent tx, preview, apply, and undo."""
    agent = DocumentAgent(sample_docx)
    blk = agent.canonical_doc.sections[0].blocks[0]
    
    # Build selection context
    ctx = agent.get_selection_context(block_id=blk.id, start=0, end=10)
    assert ctx.block_id == blk.id
    assert ctx.document_profile == "academic_report"

    # Agent proposes replacement
    op = ReplaceTextOp(
        block_id=blk.id,
        target_substring="Nghiên Cứu",
        replacement_text="Nghiên Cứu Chuyên Sâu",
    )
    preview = agent.propose_agent_transaction(
        description="Rewrite section academically",
        operations=[op],
    )
    assert preview.operations_count == 1
    assert preview.transaction_id in agent.tx_manager.pending_transactions

    # Apply transaction
    agent.apply_agent_transaction(preview.transaction_id)
    assert "Nghiên Cứu Chuyên Sâu" in blk.full_text

    # Single-click undo
    agent.undo()
    assert "Nghiên Cứu" in blk.full_text


def test_scenario_c_clarification_ambiguity():
    """SCENARIO C: Agent assesses ambiguous vs specific prompts."""
    # Ambiguous prompt -> LOW confidence -> returns multiple-choice options
    conf_low, req_low = ClarificationEngine.assess_instruction("Viết lại đoạn này")
    assert conf_low == AgentConfidence.LOW
    assert req_low is not None
    assert len(req_low.options) >= 3
    assert any(opt.id == "academic" for opt in req_low.options)

    # Specific prompt -> HIGH confidence -> proceeds directly
    conf_high, req_high = ClarificationEngine.assess_instruction("Căn đều 2 lề và giãn dòng 1.5 cho đoạn p_0001")
    assert conf_high == AgentConfidence.HIGH
    assert req_high is None


def test_scenario_d_research_citations_no_hallucination():
    """SCENARIO D: Research assistant finds verified sources and formats citations without fabrication."""
    research = ResearchAssistant()

    # Search for foundational Transformer paper
    proposal = research.evaluate_claim_and_propose_citation("Attention Is All You Need Transformer", style="apa")
    assert proposal.unsupported_warning is None
    assert "Vaswani" in proposal.proposed_intext_citation
    assert "NeurIPS" in proposal.proposed_bibliography_entry

    # Search for unsupported non-existent claim
    bad_proposal = research.evaluate_claim_and_propose_citation("NonExistentFakeTheoryXYZ12345")
    assert bad_proposal.unsupported_warning is not None
    assert "Do not fabricate" in bad_proposal.unsupported_warning


def test_scenario_f_diagram_generation():
    """SCENARIO F: Synthesizes structured Mermaid and SVG diagrams."""
    components = ["Web Client UI", "API Gateway", "Agent Orchestrator", "Document Engine", "Vector Store"]
    diag = DiagramSynthesizer.generate_architecture_diagram(components, title="Docx-Agent V2 Architecture")

    assert diag.diagram_type == "architecture"
    assert "graph TB" in diag.source_code
    assert "<svg" in diag.rendered_svg
    assert "API Gateway" in diag.rendered_svg


def test_scenario_g_large_document_modification():
    """SCENARIO G: Modify specific block in 100+ block document without corrupting unrelated blocks."""
    doc = DocumentNode(title="Large 100-Block Doc")
    doc.sections[0].blocks = [
        ParagraphBlock(runs=[RunNode(text=f"Paragraph content for block #{i}")])
        for i in range(120)
    ]

    target_blk = doc.sections[0].blocks[50]
    target_id = target_blk.id

    op = InsertTextOp(block_id=target_id, offset=0, text="[MODIFIED-PAGE-50] ")
    tx = AgentTransactionManager(doc)
    tx.execute_operation(op)

    assert "[MODIFIED-PAGE-50]" in target_blk.full_text
    # Unrelated blocks remain intact
    assert doc.sections[0].blocks[0].full_text == "Paragraph content for block #0"
    assert doc.sections[0].blocks[119].full_text == "Paragraph content for block #119"


def test_scenario_h_crash_recovery(tmp_path):
    """SCENARIO H: Records operation log and recovers unsaved snapshot upon crash."""
    ws = WorkspacePersistence(workspace_dir=tmp_path / "workspace")
    doc = DocumentNode(title="Crash Recovery Doc", version=42)
    doc.sections[0].blocks.append(ParagraphBlock(runs=[RunNode(text="Critical Unsaved Research Data")]))

    # Save snapshot
    snap_p = ws.save_snapshot(doc)
    assert snap_p.exists()
    assert ws.has_recoverable_session()

    # Simulate restart and recovery
    recovered_doc = ws.load_latest_snapshot()
    assert recovered_doc is not None
    assert recovered_doc.version == 42
    assert "Critical Unsaved Research Data" in recovered_doc.sections[0].blocks[-1].full_text


def test_scenario_i_agent_multi_op_transaction_undo():
    """SCENARIO I: 5 operations bundled into 1 Agent Transaction undone in a single call."""
    doc = DocumentNode()
    blk = ParagraphBlock(runs=[RunNode(text="Original baseline text.")])
    doc.sections[0].blocks = [blk]
    tx_manager = AgentTransactionManager(doc)

    ops = [
        InsertTextOp(block_id=blk.id, offset=0, text="[Step 1] "),
        InsertTextOp(block_id=blk.id, offset=0, text="[Step 2] "),
        InsertTextOp(block_id=blk.id, offset=0, text="[Step 3] "),
        FormatParagraphOp(block_id=blk.id, line_spacing=2.0),
        FormatParagraphOp(block_id=blk.id, alignment="right"),
    ]

    composite = CompositeOperation(description="5-step refactor", operations=ops)
    tx_manager.execute_operation(composite)

    assert "[Step 3] [Step 2] [Step 1]" in blk.full_text
    assert blk.line_spacing == 2.0
    assert blk.alignment == "right"

    # Single undo reverts all 5 operations
    tx_manager.undo()

    assert blk.full_text == "Original baseline text."
    assert blk.line_spacing == 1.5
    assert blk.alignment == "justify"


def test_scenario_j_unsupported_element_preservation():
    """SCENARIO J: Preserves unsupported element tag without silent loss."""
    raw_xml_data = "<w:drawing xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:customShape/></w:drawing>"
    unsupported = UnsupportedBlock(tag_name="w:drawing", raw_xml=raw_xml_data)

    doc = DocumentNode()
    doc.sections[0].blocks.append(unsupported)

    found = doc.find_block(unsupported.id)
    assert found is not None
    assert isinstance(found, UnsupportedBlock)
    assert "customShape" in found.raw_xml


def test_visual_layout_verification():
    """Validates VisualLayoutVerifier catches overflows and heading skips."""
    doc = DocumentNode()
    sec = doc.sections[0]

    # Heading skip: H1 followed by H3
    h1 = HeadingBlock(level=1, runs=[RunNode(text="Chapter 1")])
    h3 = HeadingBlock(level=3, runs=[RunNode(text="Section 1.1.1")])
    
    # Image overflow: 25cm width on A4 page (printable width = 16cm)
    img_overflow = ImageBlock(width_cm=25.0)

    sec.blocks = [h1, h3, img_overflow]

    report = VisualLayoutVerifier.verify_document_layout(doc)
    assert not report.passed
    anomaly_types = [a.anomaly_type for a in report.anomalies]
    assert "HEADING_HIERARCHY" in anomaly_types
    assert "IMAGE_OVERFLOW" in anomaly_types
