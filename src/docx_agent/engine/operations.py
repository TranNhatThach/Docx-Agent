"""
Deterministic, reversible Document Operation Engine for Canonical Document Model.
"""

from abc import ABC, abstractmethod
from typing import Dict, Any, List, Optional, Union, Tuple
from pydantic import BaseModel, Field
from copy import deepcopy

from docx_agent.canonical.model import (
    DocumentNode,
    SectionNode,
    BaseBlockNode,
    ParagraphBlock,
    HeadingBlock,
    ListItemBlock,
    TableBlock,
    TableCellNode,
    ImageBlock,
    DiagramBlock,
    RunNode,
    CitationNode,
    SourceMetadata,
    BlockNode,
    generate_id,
)
from docx_agent.core.exceptions import DocxAgentError, ErrorCode


class DocOperation(BaseModel, ABC):
    """Abstract base class for all deterministic document mutations."""
    op_id: str = Field(default_factory=lambda: generate_id("op"))
    op_type: str = "base"

    @abstractmethod
    def apply(self, doc: DocumentNode) -> bool:
        """Executes operation on the document model."""
        pass

    @abstractmethod
    def invert(self) -> "DocOperation":
        """Returns the exact inverse operation to reverse this mutation."""
        pass


class InsertTextOp(DocOperation):
    op_type: str = "insert_text"
    block_id: str
    offset: int = 0
    text: str
    font_name: Optional[str] = None
    font_size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    color_rgb: Optional[str] = None

    def apply(self, doc: DocumentNode) -> bool:
        blk = doc.find_block(self.block_id)
        if not isinstance(blk, ParagraphBlock):
            return False

        blk.dirty = True
        if not blk.runs:
            blk.runs.append(
                RunNode(
                    text=self.text,
                    font_name=self.font_name,
                    font_size_pt=self.font_size_pt,
                    bold=self.bold,
                    italic=self.italic,
                    color_rgb=self.color_rgb,
                )
            )
            return True

        # Find target run by character offset
        current_offset = 0
        inserted = False
        for idx, r in enumerate(blk.runs):
            r_len = len(r.text)
            if current_offset <= self.offset <= current_offset + r_len:
                rel_offset = self.offset - current_offset
                prefix = r.text[:rel_offset]
                suffix = r.text[rel_offset:]
                
                # If formatting matches current run, insert in-place
                if (
                    self.font_name is None
                    and self.font_size_pt is None
                    and self.bold is None
                    and self.italic is None
                ):
                    r.text = prefix + self.text + suffix
                else:
                    # Split run into prefix, new formatted run, suffix
                    new_runs = []
                    if prefix:
                        r_pref = deepcopy(r)
                        r_pref.text = prefix
                        new_runs.append(r_pref)
                    new_runs.append(
                        RunNode(
                            text=self.text,
                            font_name=self.font_name or r.font_name,
                            font_size_pt=self.font_size_pt or r.font_size_pt,
                            bold=self.bold if self.bold is not None else r.bold,
                            italic=self.italic if self.italic is not None else r.italic,
                            color_rgb=self.color_rgb or r.color_rgb,
                        )
                    )
                    if suffix:
                        r_suff = deepcopy(r)
                        r_suff.text = suffix
                        new_runs.append(r_suff)
                    
                    blk.runs[idx : idx + 1] = new_runs

                inserted = True
                break
            current_offset += r_len

        if not inserted:
            blk.runs.append(RunNode(text=self.text))
        return True

    def invert(self) -> DocOperation:
        return DeleteTextOp(
            block_id=self.block_id,
            offset=self.offset,
            length=len(self.text),
        )


class DeleteTextOp(DocOperation):
    op_type: str = "delete_text"
    block_id: str
    offset: int = 0
    length: int = 0
    deleted_runs_backup: List[RunNode] = Field(default_factory=list)

    def apply(self, doc: DocumentNode) -> bool:
        blk = doc.find_block(self.block_id)
        if not isinstance(blk, ParagraphBlock):
            return False

        blk.dirty = True
        full_text = blk.full_text
        if self.offset >= len(full_text):
            return False

        target_end = min(self.offset + self.length, len(full_text))
        current_offset = 0
        new_runs: List[RunNode] = []
        self.deleted_runs_backup = []

        for r in blk.runs:
            r_start = current_offset
            r_end = current_offset + len(r.text)

            if r_end <= self.offset or r_start >= target_end:
                # Outside deletion zone
                new_runs.append(r)
            else:
                # Intersects deletion zone
                del_sub_start = max(0, self.offset - r_start)
                del_sub_end = min(len(r.text), target_end - r_start)

                del_run = deepcopy(r)
                del_run.text = r.text[del_sub_start:del_sub_end]
                self.deleted_runs_backup.append(del_run)

                surviving_prefix = r.text[:del_sub_start]
                surviving_suffix = r.text[del_sub_end:]

                if surviving_prefix:
                    r_p = deepcopy(r)
                    r_p.text = surviving_prefix
                    new_runs.append(r_p)
                if surviving_suffix:
                    r_s = deepcopy(r)
                    r_s.text = surviving_suffix
                    new_runs.append(r_s)

            current_offset = r_end

        blk.runs = new_runs if new_runs else [RunNode(text="")]
        return True

    def invert(self) -> DocOperation:
        # Reconstruct text from deleted runs
        del_text = "".join(r.text for r in self.deleted_runs_backup)
        return InsertTextOp(
            block_id=self.block_id,
            offset=self.offset,
            text=del_text,
        )


class ReplaceTextOp(DocOperation):
    op_type: str = "replace_text"
    block_id: str
    target_substring: str
    replacement_text: str
    old_runs_backup: List[RunNode] = Field(default_factory=list)

    def apply(self, doc: DocumentNode) -> bool:
        blk = doc.find_block(self.block_id)
        if not isinstance(blk, ParagraphBlock):
            return False

        self.old_runs_backup = deepcopy(blk.runs)
        blk.dirty = True

        full_text = blk.full_text
        match_idx = full_text.find(self.target_substring)
        if match_idx == -1:
            return False

        # Execute Delete + Insert via run segment splitting
        del_op = DeleteTextOp(
            block_id=self.block_id,
            offset=match_idx,
            length=len(self.target_substring),
        )
        del_op.apply(doc)

        ins_op = InsertTextOp(
            block_id=self.block_id,
            offset=match_idx,
            text=self.replacement_text,
        )
        ins_op.apply(doc)
        return True

    def invert(self) -> DocOperation:
        return RestoreBlockRunsOp(
            block_id=self.block_id,
            runs=self.old_runs_backup,
        )


class RestoreBlockRunsOp(DocOperation):
    op_type: str = "restore_block_runs"
    block_id: str
    runs: List[RunNode]

    def apply(self, doc: DocumentNode) -> bool:
        blk = doc.find_block(self.block_id)
        if isinstance(blk, ParagraphBlock):
            blk.runs = deepcopy(self.runs)
            blk.dirty = True
            return True
        return False

    def invert(self) -> DocOperation:
        return RestoreBlockRunsOp(block_id=self.block_id, runs=self.runs)


class FormatParagraphOp(DocOperation):
    op_type: str = "format_paragraph"
    block_id: str
    alignment: Optional[str] = None
    line_spacing: Optional[float] = None
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None
    first_line_indent_cm: Optional[float] = None
    old_state: Dict[str, Any] = Field(default_factory=dict)

    def apply(self, doc: DocumentNode) -> bool:
        blk = doc.find_block(self.block_id)
        if not isinstance(blk, ParagraphBlock):
            return False

        self.old_state = {
            "alignment": blk.alignment,
            "line_spacing": blk.line_spacing,
            "space_before_pt": blk.space_before_pt,
            "space_after_pt": blk.space_after_pt,
            "first_line_indent_cm": blk.first_line_indent_cm,
        }

        if self.alignment is not None:
            blk.alignment = self.alignment
        if self.line_spacing is not None:
            blk.line_spacing = self.line_spacing
        if self.space_before_pt is not None:
            blk.space_before_pt = self.space_before_pt
        if self.space_after_pt is not None:
            blk.space_after_pt = self.space_after_pt
        if self.first_line_indent_cm is not None:
            blk.first_line_indent_cm = self.first_line_indent_cm

        blk.dirty = True
        return True

    def invert(self) -> DocOperation:
        return FormatParagraphOp(
            block_id=self.block_id,
            alignment=self.old_state.get("alignment"),
            line_spacing=self.old_state.get("line_spacing"),
            space_before_pt=self.old_state.get("space_before_pt"),
            space_after_pt=self.old_state.get("space_after_pt"),
            first_line_indent_cm=self.old_state.get("first_line_indent_cm"),
        )


class InsertBlockOp(DocOperation):
    op_type: str = "insert_block"
    section_index: int = 0
    block_index: int = 0
    block: BlockNode

    def apply(self, doc: DocumentNode) -> bool:
        if 0 <= self.section_index < len(doc.sections):
            sec = doc.sections[self.section_index]
            idx = max(0, min(self.block_index, len(sec.blocks)))
            sec.blocks.insert(idx, deepcopy(self.block))
            return True
        return False

    def invert(self) -> DocOperation:
        return DeleteBlockOp(
            section_index=self.section_index,
            block_id=self.block.id,
            deleted_block_backup=deepcopy(self.block),
            block_index=self.block_index,
        )


class DeleteBlockOp(DocOperation):
    op_type: str = "delete_block"
    section_index: int = 0
    block_id: str
    block_index: int = 0
    deleted_block_backup: Optional[BlockNode] = None

    def apply(self, doc: DocumentNode) -> bool:
        if 0 <= self.section_index < len(doc.sections):
            sec = doc.sections[self.section_index]
            for idx, b in enumerate(sec.blocks):
                if b.id == self.block_id:
                    self.block_index = idx
                    self.deleted_block_backup = deepcopy(b)
                    sec.blocks.pop(idx)
                    return True
        return False

    def invert(self) -> DocOperation:
        if self.deleted_block_backup:
            return InsertBlockOp(
                section_index=self.section_index,
                block_index=self.block_index,
                block=self.deleted_block_backup,
            )
        raise DocxAgentError("Cannot invert DeleteBlockOp without deleted_block_backup", ErrorCode.TRANSACTION_FAILED)


class UpdateTableCellOp(DocOperation):
    op_type: str = "update_table_cell"
    table_id: str
    row: int
    col: int
    text: str
    bold: Optional[bool] = None
    bg_color_hex: Optional[str] = None
    old_text: str = ""
    old_bg: Optional[str] = None

    def apply(self, doc: DocumentNode) -> bool:
        blk = doc.find_block(self.table_id)
        if not isinstance(blk, TableBlock):
            return False

        if 0 <= self.row < len(blk.cells) and 0 <= self.col < len(blk.cells[self.row]):
            cell = blk.cells[self.row][self.col]
            self.old_text = cell.text
            self.old_bg = cell.bg_color_hex

            cell.text = self.text
            cell.runs = [RunNode(text=self.text, bold=self.bold)]
            if self.bg_color_hex:
                cell.bg_color_hex = self.bg_color_hex
            blk.dirty = True
            return True
        return False

    def invert(self) -> DocOperation:
        return UpdateTableCellOp(
            table_id=self.table_id,
            row=self.row,
            col=self.col,
            text=self.old_text,
            bg_color_hex=self.old_bg,
        )


class InsertCitationOp(DocOperation):
    op_type: str = "insert_citation"
    block_id: str
    offset: int
    source: SourceMetadata
    citation_style: str = "apa"

    def apply(self, doc: DocumentNode) -> bool:
        # Register source in document
        doc.sources[self.source.id] = self.source

        # Format citation text
        intext = f"({self.source.authors[0] if self.source.authors else 'Anon'}, {self.source.year or 'n.d.'})"
        if self.citation_style == "ieee":
            cite_num = len(doc.citations) + 1
            intext = f"[{cite_num}]"

        cite_node = CitationNode(
            source_id=self.source.id,
            citation_style=self.citation_style,
            formatted_intext=intext,
        )
        doc.citations[cite_node.id] = cite_node

        # Insert citation run
        ins = InsertTextOp(
            block_id=self.block_id,
            offset=self.offset,
            text=f" {intext}",
            font_size_pt=10.0 if self.citation_style == "ieee" else None,
            superscript=True if self.citation_style == "ieee" else None,
        )
        return ins.apply(doc)

    def invert(self) -> DocOperation:
        # Revert text insertion
        intext = f"({self.source.authors[0] if self.source.authors else 'Anon'}, {self.source.year or 'n.d.'})"
        if self.citation_style == "ieee":
            intext = f"[{len(self.source.title)}]"
        return DeleteTextOp(
            block_id=self.block_id,
            offset=self.offset,
            length=len(intext) + 1,
        )


class CompositeOperation(DocOperation):
    """
    Groups multiple operations into a single atomic transaction.
    """
    op_type: str = "composite"
    transaction_id: str = Field(default_factory=lambda: generate_id("tx"))
    description: str = "Agent Transaction"
    operations: List[DocOperation] = Field(default_factory=list)

    def apply(self, doc: DocumentNode) -> bool:
        applied_ops = []
        for op in self.operations:
            success = op.apply(doc)
            if not success:
                # Rollback previously applied sub-operations
                for applied in reversed(applied_ops):
                    applied.invert().apply(doc)
                return False
            applied_ops.append(op)
        return True

    def invert(self) -> DocOperation:
        inverse_ops = [op.invert() for op in reversed(self.operations)]
        return CompositeOperation(
            transaction_id=f"{self.transaction_id}_undo",
            description=f"Undo: {self.description}",
            operations=inverse_ops,
        )
