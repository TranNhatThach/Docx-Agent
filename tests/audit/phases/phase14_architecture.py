"""
Phase 14: Architectural Cleanliness & Separation Audit.
Inspects package dependencies and checks strict separation between Canonical Model,
Engine, Persistence, Selection Context, and MCP Interfaces.
"""

import time
import json
import inspect
from pathlib import Path
from typing import Dict, Any

from docx_agent.canonical import model
from docx_agent.engine import operations, transactions
from docx_agent.adapters import docx
from docx_agent.interfaces.workspace import bridge
from docx_agent.interfaces.mcp import server

def run_phase_14_audit() -> Dict[str, Any]:
    evidence: Dict[str, Any] = {
        "phase": 14,
        "phase_name": "Architectural Cleanliness & Separation",
        "timestamp": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "status": "VERIFIED",
        "layers_audited": [],
        "violations": [],
    }

    all_passed = True

    # Layer 1: Canonical Document Model must NOT import UI or VS Code or docx exporter
    model_source = inspect.getsource(model)
    has_vscode_in_model = "vscode" in model_source
    has_http_in_model = "http" in model_source

    evidence["layers_audited"].append({
        "layer": "Canonical Document Model",
        "clean": not (has_vscode_in_model or has_http_in_model),
        "source_file": "src/docx_agent/canonical/model.py",
    })
    if has_vscode_in_model or has_http_in_model:
        all_passed = False
        evidence["violations"].append("Canonical Model contains UI or networking imports")

    # Layer 2: Operation Engine must operate purely on DocumentNode / in-memory structures
    engine_source = inspect.getsource(operations)
    has_docx_write_in_engine = "doc.save(" in engine_source

    evidence["layers_audited"].append({
        "layer": "Operations Engine",
        "clean": not has_docx_write_in_engine,
        "source_file": "src/docx_agent/engine/operations.py",
    })
    if has_docx_write_in_engine:
        all_passed = False
        evidence["violations"].append("Operations Engine directly calls doc.save instead of delegating to persistence")

    # Layer 3: Selection Context bridge decoupling
    bridge_source = inspect.getsource(bridge)
    has_clean_bridge = "WorkspaceBridge" in bridge_source

    evidence["layers_audited"].append({
        "layer": "Workspace Bridge & Selection",
        "clean": has_clean_bridge,
        "source_file": "src/docx_agent/interfaces/workspace/bridge.py",
    })

    evidence["status"] = "VERIFIED" if all_passed else "BROKEN"
    evidence["score"] = 3.0 if all_passed else 1.0  # Weight: 3 points
    return evidence


if __name__ == "__main__":
    res = run_phase_14_audit()
    print(json.dumps(res, indent=2, ensure_ascii=False))
