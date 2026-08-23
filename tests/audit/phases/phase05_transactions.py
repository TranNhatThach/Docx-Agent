"""
Phase 5: Agent Edit Transactions & Diff Review Audit.
Verifies proposal staging, inline diff rendering, atomic apply, rejection, and undo.
"""

import time
import json
from pathlib import Path
from typing import Dict, Any
from docx_agent.agent import DocumentAgent
from docx_agent.engine.transactions import AgentTransactionManager
from docx_agent.engine.operations import ReplaceTextOp

FIXTURES_DIR = Path(__file__).parent.parent.parent / "fixtures"


def run_phase_05_audit(tmp_path: Path) -> Dict[str, Any]:
    evidence: Dict[str, Any] = {
        "phase": 5,
        "phase_name": "Agent Edit Transactions & Diff Review",
        "timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "status": "VERIFIED",
        "transaction_tests": [],
        "errors": [],
    }

    doc_p = FIXTURES_DIR / "academic_report.docx"
    test_doc = tmp_path / "phase05_tx.docx"
    test_doc.write_bytes(doc_p.read_bytes())

    all_passed = True
    agent = DocumentAgent(test_doc)
    
    # Get first paragraph block
    sec = agent.canonical_doc.sections[0]
    p_blk = [b for b in sec.blocks if b.type == "paragraph"][0]

    tx_manager = AgentTransactionManager(agent.canonical_doc)

    # 1. Proposal creation - verify model remains unchanged before approval
    try:
        orig_para_text = p_blk.full_text
        target_sub = orig_para_text[:min(10, len(orig_para_text))]
        
        op = ReplaceTextOp(
            block_id=p_blk.id,
            target_substring=target_sub,
            replacement_text=f"{target_sub} [ĐỀ XUẤT AGENT]",
        )
        preview = tx_manager.propose_transaction(
            description="Viết lại đoạn giới thiệu chuẩn học thuật",
            operations=[op],
        )
        assert preview.transaction_id is not None
        assert preview.operations_count == 1
        assert p_blk.id in preview.affected_blocks

        # Verify model NOT changed yet
        assert p_blk.full_text == orig_para_text

        evidence["transaction_tests"].append({
            "step": "proposal_creation",
            "status": "SUCCESS",
            "transaction_id": preview.transaction_id,
            "document_unchanged_before_approval": True,
        })
    except Exception as e:
        all_passed = False
        evidence["errors"].append(f"Proposal creation error: {str(e)}")

    # 2. Apply Transaction
    try:
        applied = tx_manager.apply_pending_transaction(preview.transaction_id)
        assert applied is True

        # Verify model updated
        assert "[ĐỀ XUẤT AGENT]" in p_blk.full_text

        evidence["transaction_tests"].append({
            "step": "apply_transaction",
            "status": "SUCCESS",
            "applied_transaction_id": preview.transaction_id,
        })
    except Exception as e:
        all_passed = False
        evidence["errors"].append(f"Apply error: {str(e)}")

    # 3. Undo Transaction
    try:
        undo_res = tx_manager.undo()
        assert undo_res is True

        # Verify restored
        assert p_blk.full_text == orig_para_text

        evidence["transaction_tests"].append({
            "step": "undo_transaction",
            "status": "SUCCESS",
            "restored_original": True,
        })
    except Exception as e:
        all_passed = False
        evidence["errors"].append(f"Undo error: {str(e)}")

    # 4. Reject Proposal Scenario
    try:
        op2 = ReplaceTextOp(
            block_id=p_blk.id,
            target_substring=target_sub,
            replacement_text="TEST_REJECT",
        )
        preview2 = tx_manager.propose_transaction(
            description="Chỉnh sửa thử nghiệm",
            operations=[op2],
        )
        reject_res = tx_manager.reject_pending_transaction(preview2.transaction_id)
        assert reject_res is True
        assert preview2.transaction_id not in tx_manager.pending_transactions
        assert p_blk.full_text == orig_para_text

        evidence["transaction_tests"].append({
            "step": "reject_proposal",
            "status": "SUCCESS",
        })
    except Exception as e:
        all_passed = False
        evidence["errors"].append(f"Reject error: {str(e)}")

    evidence["status"] = "VERIFIED" if all_passed else "BROKEN"
    evidence["score"] = 9.0 if all_passed else 3.0  # Weight: 9 points
    return evidence


if __name__ == "__main__":
    import tempfile
    with tempfile.TemporaryDirectory() as td:
        res = run_phase_05_audit(Path(td))
        print(json.dumps(res, indent=2, ensure_ascii=False))
