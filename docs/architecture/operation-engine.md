# Operation & Command Engine Architecture

## 1. Core Principles
Every human edit and AI Agent proposal is expressed as a deterministic, serializable `DocOperation` adhering to:
- `apply(doc: DocumentNode) -> bool`
- `invert() -> DocOperation`

## 2. Operation Taxonomy

| Operation | Purpose | Inversion Strategy |
| :--- | :--- | :--- |
| `InsertTextOp` | Inserts text at character offset | `DeleteTextOp` with corresponding length |
| `DeleteTextOp` | Deletes text range | `InsertTextOp` with backed-up deleted runs |
| `ReplaceTextOp` | Replaces target substring | Restores backed-up runs |
| `FormatParagraphOp` | Updates spacing, alignment, indents | Restores previous paragraph format state |
| `InsertBlockOp` | Inserts block at section index | `DeleteBlockOp` with block ID |
| `DeleteBlockOp` | Deletes block | `InsertBlockOp` with backed-up block |
| `InsertCitationOp` | Inserts citation run and source metadata | Removes citation run and decrements references |
| `CompositeOperation` | Groups multiple ops into atomic transaction | Reverses inverse operations in reverse order |

## 3. Agent Transactions & Undo/Redo
The `AgentTransactionManager` supports:
- Staging proposals (`propose_transaction()`)
- Rich user previews with line diffs
- Atomic application on approval
- Single-click full transaction undo (`undo()`)
