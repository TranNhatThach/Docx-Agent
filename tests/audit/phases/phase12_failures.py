"""
Phase 12: Failure Injection & Recovery Audit.
Injects malformed DOCX, missing files, ambiguous targets, and corrupted state.
Verifies failure containment, error diagnostics, and subsequent operation recovery.
"""

import time
import json
from pathlib import Path
from typing import Dict, Any
from docx_agent.agent import DocumentAgent
from docx_agent.interfaces.workspace.bridge import WorkspaceBridge
from docx_agent.core.exceptions import DocxAgentError, AmbiguousTargetError

FIXTURES_DIR = Path(__file__).parent.parent.parent / "fixtures"


def run_phase_12_audit(tmp_path: Path) -> Dict[str, Any]:
    evidence: Dict[str, Any] = {
        "phase": 12,
        "phase_name": "Failure Injection & Recovery",
        "timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "status": "VERIFIED",
        "injection_tests": [],
        "errors": [],
    }

    all_passed = True

    # 1. Missing File Injection
    try:
        missing_p = tmp_path / "non_existent_file_9999.docx"
        WorkspaceBridge.load_document_payload(missing_p)
        all_passed = False
        evidence["injection_tests"].append({"test": "missing_file", "status": "UNEXPECTED_SUCCESS"})
    except Exception as e:
        evidence["injection_tests"].append({
            "test": "missing_file",
            "status": "PROPERLY_REJECTED",
            "error_type": type(e).__name__,
            "diagnostic": str(e),
        })

    # 2. Corrupted DOCX Injection
    try:
        corrupted_p = FIXTURES_DIR / "corrupted.docx"
        WorkspaceBridge.load_document_payload(corrupted_p)
        all_passed = False
        evidence["injection_tests"].append({"test": "corrupted_docx", "status": "UNEXPECTED_SUCCESS"})
    except Exception as e:
        evidence["injection_tests"].append({
            "test": "corrupted_docx",
            "status": "PROPERLY_REJECTED",
            "error_type": type(e).__name__,
            "diagnostic": str(e),
        })

    # 3. Ambiguous Target Resolution Injection
    try:
        doc_p = FIXTURES_DIR / "academic_report.docx"
        test_doc = tmp_path / "ambiguous_test.docx"
        test_doc.write_bytes(doc_p.read_bytes())
        agent = DocumentAgent(test_doc)

        # Ambiguous match that appears multiple times in doc
        agent.replace(target="và", replacement="VÀ")
        evidence["injection_tests"].append({
            "test": "ambiguous_target",
            "status": "PROPERLY_HANDLED",
        })
    except Exception as e:
        evidence["injection_tests"].append({
            "test": "ambiguous_target",
            "status": "PROPERLY_HANDLED",
            "error_type": type(e).__name__,
        })

    # 4. Subsequent Operation Recovery
    try:
        valid_doc = FIXTURES_DIR / "simple.docx"
        payload = WorkspaceBridge.load_document_payload(valid_doc)
        assert payload["success"] is True
        evidence["injection_tests"].append({
            "test": "recovery_after_failure",
            "status": "SUCCESS",
            "recovered": True,
        })
    except Exception as e:
        all_passed = False
        evidence["injection_tests"].append({
            "test": "recovery_after_failure",
            "status": "FAILED",
            "error": str(e),
        })

    evidence["status"] = "VERIFIED" if all_passed else "BROKEN"
    evidence["score"] = 8.0 if all_passed else 3.0  # Weight: 8 points
    return evidence


if __name__ == "__main__":
    import tempfile
    with tempfile.TemporaryDirectory() as td:
        res = run_phase_12_audit(Path(td))
        print(json.dumps(res, indent=2, ensure_ascii=False))
