"""
Agent-safe logging utility.
Separates diagnostic logs (to stderr) from clean machine-readable stdout JSON.
"""

import sys
import json
from typing import Any, Dict, Optional


class AgentLogger:
    def __init__(self, verbose: bool = False, quiet: bool = False):
        self.verbose = verbose
        self.quiet = quiet

    def info(self, msg: str) -> None:
        if not self.quiet:
            sys.stderr.write(f"[INFO] {msg}\n")
            sys.stderr.flush()

    def debug(self, msg: str) -> None:
        if self.verbose and not self.quiet:
            sys.stderr.write(f"[DEBUG] {msg}\n")
            sys.stderr.flush()

    def warning(self, msg: str) -> None:
        if not self.quiet:
            sys.stderr.write(f"[WARN] {msg}\n")
            sys.stderr.flush()

    def error(self, msg: str) -> None:
        sys.stderr.write(f"[ERROR] {msg}\n")
        sys.stderr.flush()

    def emit_json(self, data: Any) -> None:
        """Outputs clean JSON to stdout for machine consumption."""
        if hasattr(data, "model_dump"):
            dumped = data.model_dump(exclude_none=True)
        elif hasattr(data, "to_dict"):
            dumped = data.to_dict()
        else:
            dumped = data
        
        json_str = json.dumps(dumped, indent=2, ensure_ascii=False)
        sys.stdout.write(json_str + "\n")
        sys.stdout.flush()


logger = AgentLogger()
