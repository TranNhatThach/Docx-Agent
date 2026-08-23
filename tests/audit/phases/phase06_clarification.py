"""
Phase 6: Clarification Engine Audit.
Verifies material ambiguity detection, confidence scoring, and multi-choice question generation.
"""

import time
import json
from typing import Dict, Any
from docx_agent.engine.clarification import ClarificationEngine, AgentConfidence

def run_phase_06_audit() -> Dict[str, Any]:
    evidence: Dict[str, Any] = {
        "phase": 6,
        "phase_name": "Clarification Engine",
        "timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "status": "VERIFIED",
        "clarification_cases": [],
        "errors": [],
    }

    all_passed = True

    # Case 1: Ambiguous style request
    try:
        conf1, req1 = ClarificationEngine.assess_instruction(
            instruction="Viết lại đoạn này cho hay hơn",
            selected_text="Kiến trúc hệ thống bao gồm ba tầng chính.",
        )
        assert conf1 in (AgentConfidence.LOW, AgentConfidence.MEDIUM)
        assert req1 is not None
        assert len(req1.options) >= 2

        evidence["clarification_cases"].append({
            "instruction": "Viết lại đoạn này cho hay hơn",
            "confidence": conf1.value,
            "clarification_triggered": True,
            "question": req1.question,
            "options": [opt.model_dump() for opt in req1.options],
        })
    except Exception as e:
        all_passed = False
        evidence["errors"].append(f"Ambiguity case 1 error: {str(e)}")

    # Case 2: Clear, explicit formatting instruction
    try:
        conf2, req2 = ClarificationEngine.assess_instruction(
            instruction="Đổi font chữ đoạn 2 sang Times New Roman 13pt và căn đều",
            selected_text="Đoạn văn cần đổi định dạng.",
        )
        assert conf2 == AgentConfidence.HIGH
        assert req2 is None

        evidence["clarification_cases"].append({
            "instruction": "Đổi font chữ đoạn 2 sang Times New Roman 13pt và căn đều",
            "confidence": conf2.value,
            "clarification_triggered": False,
        })
    except Exception as e:
        all_passed = False
        evidence["errors"].append(f"Clear case error: {str(e)}")

    # Case 3: Diagram type ambiguity
    try:
        conf3, req3 = ClarificationEngine.assess_instruction(
            instruction="Vẽ sơ đồ mô tả hệ thống này",
            selected_text="Hệ thống gồm Frontend, Backend, Database và AI Agent.",
        )
        assert req3 is not None
        evidence["clarification_cases"].append({
            "instruction": "Vẽ sơ đồ mô tả hệ thống này",
            "confidence": conf3.value,
            "clarification_triggered": True,
            "question": req3.question,
        })
    except Exception as e:
        all_passed = False
        evidence["errors"].append(f"Diagram ambiguity case error: {str(e)}")

    evidence["status"] = "VERIFIED" if all_passed else "BROKEN"
    evidence["score"] = 5.0 if all_passed else 2.0  # Weight: 5 points
    return evidence


if __name__ == "__main__":
    res = run_phase_06_audit()
    print(json.dumps(res, indent=2, ensure_ascii=False))
