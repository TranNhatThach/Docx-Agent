"""
Backup management for safe in-place mutations.
"""

import shutil
from datetime import datetime
from pathlib import Path
from typing import Union, Optional
from docx_agent.utils.paths import resolve_safe_path
from docx_agent.core.exceptions import TransactionError


class BackupManager:
    """
    Creates and restores shadow backups prior to document mutations.
    """

    @staticmethod
    def create_backup(file_path: Union[str, Path], use_timestamp: bool = False) -> Path:
        """
        Creates a backup of the target file.
        Returns the path to the created backup file.
        """
        source = resolve_safe_path(file_path)
        if not source.exists():
            raise FileNotFoundError(f"Cannot backup non-existent file: {source}")

        if use_timestamp:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = source.with_name(f"{source.stem}_{ts}.bak.docx")
        else:
            backup_path = source.with_suffix(source.suffix + ".bak")

        try:
            shutil.copy2(source, backup_path)
            return backup_path
        except Exception as e:
            raise TransactionError(
                message=f"Failed to create backup at '{backup_path}': {str(e)}",
                step="backup",
                original_error=e,
            )

    @staticmethod
    def restore_backup(backup_path: Union[str, Path], target_path: Union[str, Path]) -> None:
        """Restores the target file from backup."""
        b_path = resolve_safe_path(backup_path)
        t_path = resolve_safe_path(target_path)
        if not b_path.exists():
            raise FileNotFoundError(f"Backup file not found: {b_path}")
        shutil.copy2(b_path, t_path)
