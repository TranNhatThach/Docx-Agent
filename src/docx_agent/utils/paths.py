"""
Safe path resolution utilities for docx-agent.
"""

import os
from pathlib import Path
from typing import Union


def resolve_safe_path(path: Union[str, Path]) -> Path:
    """
    Resolves and normalizes a file path.
    Expands user home directory and environment variables.
    """
    if isinstance(path, str):
        path = os.path.expanduser(os.path.expandvars(path))
    p = Path(path).resolve()
    return p


def ensure_parent_dir(path: Union[str, Path]) -> Path:
    """Ensures parent directory exists before writing."""
    p = resolve_safe_path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    return p
