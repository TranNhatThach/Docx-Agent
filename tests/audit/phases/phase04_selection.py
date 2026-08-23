"""
Phase 4: Selection Context & Antigravity MCP Bridge Audit.
Verifies runtime selection capture, SelectionContext serialization, and MCP live response.
"""

import time
import json
from pathlib import Path
from typing import Dict, Any
from docx_agent.interfaces.workspace.bridge import WorkspaceBridge
from docx_agent.interfaces.mcp.server import handle_tool_call

FIXTURES_DIR = Path(__file__).parent.parent.parent / "fixtures"


def run_phase_04_audit() -> Dict[str, Any]:
    evidence: Dict[str, Any] = {
        "phase": 4,
        "phase_name": "Selection Context & Antigravity MCP",
        "timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "status": "VERIFIED",
        "selection_tests": [],
        "errors": [],
    }

    doc_p = FIXTURES_DIR / "academic_report.docx"
    payload = WorkspaceBridge.load_document_payload(doc_p)

    h1_block = [h for h in payload["headings"] if h["level"] == 1][0]
    p_block = [b for b in payload["sections"][0]["blocks"] if b["type"] == "paragraph"][0]

    all_passed = True
    tmp_workspace = Path(".docx_agent_workspace")
    tmp_workspace.mkdir(exist_ok=True)
    sel_file = tmp_workspace / "active_selection.json"

    # 1. Selection on Heading
    try:
        ctx1 = WorkspaceBridge.get_selection_context_payload(
            file_path=doc_p,
            block_id=h1_block["id"],
            start_offset=0,
            end_offset=12,
        )
        assert ctx1["block_id"] == h1_block["id"]
        assert len(ctx1["selected_text"]) > 0

        # Simulate Webview -> Extension persistence
        with open(sel_file, "w", encoding="utf-8") as f:
            json.dump(ctx1, f, ensure_ascii=False)

        # Call MCP tool
        mcp_res = handle_tool_call("docx_get_current_selection", {})
        assert mcp_res["block_id"] == h1_block["id"]
        assert mcp_res["selected_text"] == ctx1["selected_text"]

        evidence["selection_tests"].append({
            "target": "heading",
            "block_id": h1_block["id"],
            "selected_text": ctx1["selected_text"],
            "mcp_verified": True,
        })
    except Exception as e:
        all_passed = False
        evidence["errors"].append(f"Heading selection error: {str(e)}")

    # 2. Selection on Paragraph (Changing Selection)
    try:
        ctx2 = WorkspaceBridge.get_selection_context_payload(
            file_path=doc_p,
            block_id=p_block["id"],
            start_offset=5,
            end_offset=25,
        )
        assert ctx2["block_id"] == p_block["id"]

        # Update persistence
        with open(sel_file, "w", encoding="utf-8") as f:
            json.dump(ctx2, f, ensure_ascii=False)

        # Call MCP tool again - must NOT return stale heading selection
        mcp_res2 = handle_tool_call("docx_get_current_selection", {})
        assert mcp_res2["block_id"] == p_block["id"]
        assert mcp_res2["selected_text"] == ctx2["selected_text"]
        assert mcp_res2["block_id"] != h1_block["id"]

        evidence["selection_tests"].append({
            "target": "paragraph",
            "block_id": p_block["id"],
            "selected_text": ctx2["selected_text"],
            "mcp_updated_verified": True,
        })
    except Exception as e:
        all_passed = False
        evidence["errors"].append(f"Paragraph selection error: {str(e)}")

    evidence["status"] = "VERIFIED" if all_passed else "BROKEN"
    evidence["score"] = 7.0 if all_passed else 2.0  # Weight: 7 points
    return evidence


if __name__ == "__main__":
    res = run_phase_04_audit()
    print(json.dumps(res, indent=2, ensure_ascii=False))
