# ADR-0004: Transactional Mutation with Atomic Rollback & Save Verification

## Status
Accepted

## Context
When AI agents perform multi-step document edits (e.g. batch search-and-replace, re-formatting academic presets, inserting citation sections), a failure mid-way could leave the `.docx` file in a corrupted or half-modified state. Furthermore, external applications (like Microsoft Word or Windows Explorer) may hold file locks.

## Decision
We implemented `TransactionContext` (`src/docx_agent/transactions/transaction.py`) and `WorkspaceBridge.save_document_payload`:
1. **Pre-mutation Snapshot**: Before any batch operation, an atomic backup snapshot (`.bak`) is created in a protected directory.
2. **Transaction Isolation**: Changes are applied in-memory to the Canonical Document Model first.
3. **Atomic Write & Integrity Verification**: The file is written atomically, re-opened, and validated with `DocumentValidator`. If validation fails or an exception occurs, the transaction automatically rolls back to the backup snapshot.
4. **File-Lock Detection**: If an external program locks the file, `docx-agent` catches `PermissionError`, safely saves to an alternate path (e.g. `*_Chuan_UTC.docx`), and informs the user with actionable instructions.

## Consequences
### Positive:
- Zero data corruption guarantee for agent-driven workflows.
- Full undo/redo capability across multi-step agent batches.
- Clear diagnostics and non-destructive recovery when files are locked by Microsoft Word.

### Trade-offs:
- Temporary disk usage for `.bak` files during active transaction windows (automatically cleaned up upon commit).
