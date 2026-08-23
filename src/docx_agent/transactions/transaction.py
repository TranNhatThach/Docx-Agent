"""
Transaction Manager: Enforces atomic mutation staging, verification, and automatic rollback.
"""

import os
import shutil
import tempfile
from pathlib import Path
from typing import Callable, Any, Optional, Dict, Union
import docx
from docx_agent.core.document import DocumentModel
from docx_agent.core.exceptions import TransactionError, VerificationError
from docx_agent.transactions.backup import BackupManager
from docx_agent.verification.validator import DocumentValidator
from docx_agent.utils.paths import resolve_safe_path


class TransactionContext:
    """
    Executes operations within an atomic sandbox.
    If any error occurs or verification fails, rolls back to original state.
    """

    def __init__(
        self,
        file_path: Union[str, Path],
        auto_backup: bool = True,
        verify_on_commit: bool = True,
    ):
        self.file_path = resolve_safe_path(file_path)
        self.auto_backup = auto_backup
        self.verify_on_commit = verify_on_commit
        self.backup_path: Optional[Path] = None
        self.temp_file: Optional[Path] = None
        self.model: Optional[DocumentModel] = None

    def __enter__(self) -> "TransactionContext":
        if self.file_path.exists():
            # 1. Validation of source file
            report = DocumentValidator.verify_integrity(self.file_path)
            if not report.passed:
                raise TransactionError(
                    message=f"Source document failed initial integrity check: {report.failures}",
                    step="validate_source",
                )

            # 2. Backup
            if self.auto_backup:
                self.backup_path = BackupManager.create_backup(self.file_path)

            # 3. Load into model
            self.model = DocumentModel(self.file_path)
        else:
            # Creating a new document
            self.model = DocumentModel(docx.Document())
            self.model.file_path = str(self.file_path)

        return self

    def save_and_verify(self) -> None:
        """
        Saves document to temporary staging file, independently reloads,
        runs validation, and atomically commits to the target path.
        """
        if self.model is None:
            raise TransactionError("Transaction context uninitialized", step="save")

        # 4. Save to Staging Temp File
        fd, temp_path_str = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        self.temp_file = Path(temp_path_str)

        try:
            self.model.doc.save(str(self.temp_file))
        except Exception as e:
            self.rollback()
            raise TransactionError(f"Failed to save temporary staging document: {str(e)}", step="save_temp", original_error=e)

        # 5. Independent Verification on Temp File
        if self.verify_on_commit:
            v_report = DocumentValidator.verify_integrity(self.temp_file)
            if not v_report.passed:
                self.rollback()
                raise VerificationError(
                    message=f"Independent validation failed on staging document: {v_report.failures}",
                    failures=v_report.failures,
                )

        # 6. Atomic Commit: Replace original file with verified temp file
        try:
            shutil.copy2(self.temp_file, self.file_path)
        except Exception as e:
            self.rollback()
            raise TransactionError(f"Failed to commit verified document: {str(e)}", step="commit", original_error=e)
        finally:
            if self.temp_file and self.temp_file.exists():
                try:
                    self.temp_file.unlink()
                except Exception:
                    pass

    def rollback(self) -> None:
        """Restores the original file from backup if modified and cleans up staging files."""
        if self.backup_path and self.backup_path.exists():
            try:
                BackupManager.restore_backup(self.backup_path, self.file_path)
            except Exception:
                pass

        if self.temp_file and self.temp_file.exists():
            try:
                self.temp_file.unlink()
            except Exception:
                pass

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is not None:
            # Exception occurred during context
            self.rollback()
            return False  # Re-raise exception
        return True
