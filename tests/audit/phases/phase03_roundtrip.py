"""
Phase 3: Real Save / Reopen Roundtrip Audit.
Verifies atomic save pipeline, package-level XML integrity, and semantic preservation.
"""

import time
import json
import zipfile
from pathlib import Path
from typing import Dict, Any
from docx_agent.agent import DocumentAgent
from docx_agent.interfaces.workspace.bridge import WorkspaceBridge
from docx_agent.verification.validator import DocumentValidator

FIXTURES_DIR = Path(__file__).parent.parent.parent / "fixtures"


def run_phase_03_audit(tmp_path: Path) -> Dict[str, Any]:
    evidence: Dict[str, Any] = {
        "phase": 3,
        "phase_name": "Real Save / Reopen Roundtrip",
        "timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "status": "VERIFIED",
        "roundtrip_tests": [],
        "errors": [],
    }

    test_fixtures = [
        "Vietnamese.docx",
        "academic_report.docx",
        "tables.docx",
        "images.docx",
        "multi_page.docx",
        "complex_formatting.docx",
        "unsupported_ooxml.docx",
    ]

    all_passed = True

    for fname in test_fixtures:
        src_path = FIXTURES_DIR / fname
        test_doc = tmp_path / f"rt_{fname}"
        test_doc.write_bytes(src_path.read_bytes())

        item_evidence: Dict[str, Any] = {
            "fixture": fname,
            "save_pipeline": {},
            "xml_inspection": {},
        }

        try:
            # 1. Load via Bridge
            payload = WorkspaceBridge.load_document_payload(test_doc)
            assert payload["success"] is True

            # 2. Mutate first block text with Vietnamese diacritics
            sec = payload["sections"][0]
            orig_text = sec["blocks"][0]["text"]
            sec["blocks"][0]["text"] = orig_text + " [ĐÃ KIỂM ĐỊNH ROUNDTRIP]"

            # 3. Save via Bridge (Atomic Save Pipeline)
            t0 = time.time()
            save_res = WorkspaceBridge.save_document_payload(
                file_path=test_doc,
                document_data=payload,
                output_path=test_doc,
            )
            save_latency_ms = round((time.time() - t0) * 1000, 2)

            assert save_res["success"] is True
            assert save_res["verification_passed"] is True
            item_evidence["save_pipeline"] = {
                "success": True,
                "latency_ms": save_latency_ms,
                "verification_passed": True,
            }

            # 4. Inspect ZIP Package Internals
            with zipfile.ZipFile(test_doc, "r") as z:
                namelist = z.namelist()
                item_evidence["xml_inspection"]["has_document_xml"] = "word/document.xml" in namelist
                item_evidence["xml_inspection"]["has_styles_xml"] = "word/styles.xml" in namelist
                item_evidence["xml_inspection"]["has_rels"] = "_rels/.rels" in namelist
                doc_xml = z.read("word/document.xml").decode("utf-8")
                item_evidence["xml_inspection"]["has_modified_marker"] = "ROUNDTRIP" in doc_xml

            # 5. Independent Reopen in fresh DocumentAgent
            reopened = DocumentAgent(test_doc)
            summary = reopened.inspect()
            assert summary.paragraphs_count >= 1

            # 6. Verify modified text persisted
            all_texts = [p.text for p in reopened.read()]
            assert any("[ĐÃ KIỂM ĐỊNH ROUNDTRIP]" in t for t in all_texts)
            item_evidence["status"] = "SUCCESS"

        except Exception as e:
            item_evidence["status"] = "FAILED"
            item_evidence["error"] = str(e)
            all_passed = False
            evidence["errors"].append(f"{fname}: {str(e)}")

        evidence["roundtrip_tests"].append(item_evidence)

    evidence["status"] = "VERIFIED" if all_passed else "BROKEN"
    evidence["score"] = 12.0 if all_passed else 5.0  # Weight: 12 points
    return evidence


if __name__ == "__main__":
    import tempfile
    with tempfile.TemporaryDirectory() as td:
        res = run_phase_03_audit(Path(td))
        print(json.dumps(res, indent=2, ensure_ascii=False))
