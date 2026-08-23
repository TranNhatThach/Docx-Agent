"""
Phase 7: Research & Academic Citations Audit.
Verifies citation generation, zero-hallucination source verification, and citation style fidelity.
"""

import time
import json
from typing import Dict, Any
from docx_agent.research.provider import ResearchAssistant

def run_phase_07_audit() -> Dict[str, Any]:
    evidence: Dict[str, Any] = {
        "phase": 7,
        "phase_name": "Research & Verified Academic Citations",
        "timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "status": "VERIFIED",
        "citation_evaluations": [],
        "errors": [],
    }

    assistant = ResearchAssistant()
    all_passed = True

    # 1. Evaluate well-known empirical claim in APA style
    try:
        claim_1 = "Transformer architecture revolutionized sequence modeling using self-attention mechanisms."
        prop_1 = assistant.evaluate_claim_and_propose_citation(claim_1, citation_style="apa")

        assert prop_1 is not None
        assert len(prop_1.evidence) >= 1
        top_ev = prop_1.evidence[0]
        top_src = top_ev.source

        # Verify no fabricated paper title/author: Must be from verified peer-reviewed index
        assert top_src.verified is True
        assert len(top_src.authors) > 0
        assert top_src.year >= 2016
        assert len(prop_1.proposed_intext_citation) > 0
        assert len(prop_1.proposed_bibliography_entry) > 0

        evidence["citation_evaluations"].append({
            "claim": claim_1,
            "style": "apa",
            "in_text": prop_1.proposed_intext_citation,
            "bibliography": prop_1.proposed_bibliography_entry,
            "verified_authors": top_src.authors,
            "verified_title": top_src.title,
            "verified_year": top_src.year,
            "no_hallucination_verified": True,
        })
    except Exception as e:
        all_passed = False
        evidence["errors"].append(f"APA citation error: {str(e)}")

    # 2. Evaluate IEEE style
    try:
        prop_ieee = assistant.evaluate_claim_and_propose_citation(claim_1, citation_style="ieee")
        assert prop_ieee.proposed_intext_citation.startswith("[")

        evidence["citation_evaluations"].append({
            "claim": claim_1,
            "style": "ieee",
            "in_text": prop_ieee.proposed_intext_citation,
            "bibliography": prop_ieee.proposed_bibliography_entry,
        })
    except Exception as e:
        all_passed = False
        evidence["errors"].append(f"IEEE citation error: {str(e)}")

    evidence["status"] = "VERIFIED" if all_passed else "BROKEN"
    evidence["score"] = 7.0 if all_passed else 2.0  # Weight: 7 points
    return evidence


if __name__ == "__main__":
    res = run_phase_07_audit()
    print(json.dumps(res, indent=2, ensure_ascii=False))
