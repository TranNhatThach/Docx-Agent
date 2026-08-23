"""
Agent Transaction Manager and Undo/Redo Engine for interactive workspace sessions.
"""

from typing import Dict, Any, List, Optional, Tuple
from pydantic import BaseModel, Field
from copy import deepcopy

from docx_agent.canonical.model import DocumentNode, generate_id
from docx_agent.engine.operations import DocOperation, CompositeOperation
from docx_agent.core.exceptions import DocxAgentError, ErrorCode


class TransactionPreview(BaseModel):
    transaction_id: str
    description: str
    operations_count: int
    affected_blocks: List[str]
    diff_summary: List[str]


class AgentTransactionManager:
    """
    Manages proposed Agent Transactions, user reviews, previews, atomic commits,
    and reversible multi-operation history.
    """

    def __init__(self, doc: DocumentNode):
        self.doc = doc
        self.undo_stack: List[DocOperation] = []
        self.redo_stack: List[DocOperation] = []
        self.pending_transactions: Dict[str, CompositeOperation] = {}
        self.revision: int = 1

    def propose_transaction(
        self,
        description: str,
        operations: List[DocOperation],
    ) -> TransactionPreview:
        """
        Creates and stages a proposed Agent Transaction without applying it immediately.
        Returns a rich preview for user approval.
        """
        tx = CompositeOperation(
            description=description,
            operations=operations,
        )
        self.pending_transactions[tx.transaction_id] = tx

        affected = set()
        diff_lines = []
        for op in operations:
            if hasattr(op, "block_id"):
                affected.add(getattr(op, "block_id"))
            diff_lines.append(f"{op.op_type}: {str(op.model_dump(exclude={'op_id', 'op_type'}))[:80]}")

        return TransactionPreview(
            transaction_id=tx.transaction_id,
            description=description,
            operations_count=len(operations),
            affected_blocks=sorted(list(affected)),
            diff_summary=diff_lines,
        )

    def apply_pending_transaction(self, transaction_id: str) -> bool:
        """Applies a reviewed and approved Agent Transaction."""
        tx = self.pending_transactions.pop(transaction_id, None)
        if not tx:
            raise DocxAgentError(f"Transaction not found: {transaction_id}", ErrorCode.TRANSACTION_FAILED)

        return self.execute_operation(tx)

    def reject_pending_transaction(self, transaction_id: str) -> bool:
        """Rejects and discards a pending Agent Transaction."""
        return bool(self.pending_transactions.pop(transaction_id, None))

    def execute_operation(self, op: DocOperation) -> bool:
        """Executes an operation directly (from human or approved agent) and pushes to undo stack."""
        success = op.apply(self.doc)
        if success:
            self.undo_stack.append(op)
            self.redo_stack.clear()
            self.revision += 1
            self.doc.version = self.revision
        return success

    def undo(self) -> bool:
        """Undoes the most recent operation or agent transaction."""
        if not self.undo_stack:
            return False

        op = self.undo_stack.pop()
        inverse_op = op.invert()
        success = inverse_op.apply(self.doc)
        if success:
            self.redo_stack.append(op)
            self.revision += 1
            self.doc.version = self.revision
        return success

    def redo(self) -> bool:
        """Redoes the most recently undone operation."""
        if not self.redo_stack:
            return False

        op = self.redo_stack.pop()
        success = op.apply(self.doc)
        if success:
            self.undo_stack.append(op)
            self.revision += 1
            self.doc.version = self.revision
        return success
