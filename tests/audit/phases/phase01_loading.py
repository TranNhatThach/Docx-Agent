"""
Phase 1: Real DOCX Loading & State Machine Audit.
Tests real loading across all fixtures including corrupted.docx.
Measures load, parse, and ready timings.
"""

import time
import json
from pathlib import Path
from typing import Dict, Any, List
from docx_agent.interfaces.workspace.bridge import WorkspaceBridge

FIXTURES_DIR = Path(__file__).parent.parent.parent / "fixtures"


def run_phase_01_audit() -> Dict[str, Any]:
    fixtures = [
        "simple.docx",
        "Vietnamese.docx",
        "academic_report.docx",
        "tables.docx",
        "images.docx",
        "multi_page.docx",
        "complex_formatting.docx",
        "unsupported_ooxml.docx",
        "corrupted.docx",
    ]

    evidence: Dict[str, Any] = {
        "phase": 1,
        "phase_name": "Real DOCX Loading & State Machine",
        "timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "status": "VERIFIED",
        "fixtures_tested": [],
        "errors": [],
        "metrics": {},
    }

    all_passed = True

    for fname in fixtures:
        fpath = FIXTURES_DIR / fname
        item_evidence: Dict[str, Any] = {
            "fixture": fname,
            "exists": fpath.exists(),
            "load_start": time.time(),
        }

        if fname == "corrupted.docx":
            # Expecting graceful ERROR state
            try:
                payload = WorkspaceBridge.load_document_payload(fpath)
                item_evidence["status"] = "UNEXPECTED_SUCCESS"
                all_passed = False
            except Exception as e:
                item_evidence["status"] = "ERROR_EXPECTED_AND_CAUGHT"
                item_evidence["error_type"] = type(e).__name__
                item_evidence["diagnostic"] = str(e)
                item_evidence["state_reached"] = "ERROR"
            item_evidence["load_end"] = time.time()
            item_evidence["duration_ms"] = round((item_evidence["load_end"] - item_evidence["load_start"]) * 1000, 2)
            evidence["fixtures_tested"].append(item_evidence)
            continue

        try:
            t0 = time.time()
            payload = WorkspaceBridge.load_document_payload(fpath)
            t1 = time.time()

            item_evidence["status"] = "SUCCESS"
            item_evidence["state_reached"] = "READY"
            item_evidence["duration_ms"] = round((t1 - t0) * 1000, 2)
            item_evidence["document_id"] = payload.get("document_id")
            item_evidence["title"] = payload.get("title")
            item_evidence["sections_count"] = len(payload.get("sections", []))
            item_evidence["headings_count"] = len(payload.get("headings", []))
            item_evidence["stats"] = payload.get("stats", {})

            # Inspect elements
            sec = payload["sections"][0]
            blocks = sec.get("blocks", [])
            item_evidence["blocks_count"] = len(blocks)
            item_evidence["has_headings"] = any(b.get("type") == "heading" for b in blocks)
            item_evidence["has_tables"] = any(b.get("type") == "table" for b in blocks)
            item_evidence["has_images"] = any(b.get("type") == "image" for b in blocks)
            item_evidence["has_unsupported"] = any(b.get("type") == "unsupported" for b in blocks)

        except Exception as ex:
            item_evidence["status"] = "FAILED"
            item_evidence["error"] = str(ex)
            all_passed = False
            evidence["errors"].append(f"{fname}: {str(ex)}")

        evidence["fixtures_tested"].append(item_evidence)

    evidence["status"] = "VERIFIED" if all_passed else "BROKEN"
    evidence["score"] = 6.0 if all_passed else 2.0  # Weight: 6 points
    return evidence


if __name__ == "__main__":
    res = run_phase_01_audit()
    print(json.dumps(res, indent=2, ensure_ascii=False))
