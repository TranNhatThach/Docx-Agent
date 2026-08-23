"""
Semantic Diff Engine: High-fidelity, anchor-stable document diffing algorithm.
Computes fine-grained semantic changes between DocumentNode revisions, detecting text,
paragraph, heading, table, formatting, and structural modifications with move detection.
"""

import difflib
from typing import Dict, Any, List, Optional, Tuple, Set
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
    BlockNode,
)
from docx_agent.changes.model import (
    ChangeObject,
    ChangeType,
    ChangeStatus,
    ChangeLocation,
    generate_change_id,
)
from docx_agent.utils.unicode import normalize_unicode


class SemanticDiffEngine:
    """
    Computes deterministic, reversible, semantically rich ChangeObjects between two document states.
    """

    SIMILARITY_MOVE_THRESHOLD = 0.85

    @classmethod
    def compute_diff(
        cls,
        doc_before: DocumentNode,
        doc_after: DocumentNode,
        agent_id: str = "Antigravity-Agent",
        default_reason: str = "Tối ưu hóa và chuẩn hóa nội dung tài liệu",
        confidence: float = 0.95,
        evidence: Optional[str] = None,
    ) -> List[ChangeObject]:
        changes: List[ChangeObject] = []
        doc_id = doc_before.id or doc_after.id or "doc_01"
        v_before = doc_before.version
        v_after = max(v_before + 1, doc_after.version)

        # 1. Structural Section Diffs
        changes.extend(cls._diff_sections(doc_before, doc_after, doc_id, v_before, v_after, agent_id, default_reason, confidence, evidence))

        # 2. Block-Level Semantic Matching & Diffs
        blocks_a = doc_before.all_blocks()
        blocks_b = doc_after.all_blocks()

        matched_a, matched_b, moves = cls._match_blocks(blocks_a, blocks_b)

        # Process Moved Blocks
        for idx_a, idx_b, sim in moves:
            blk_a = blocks_a[idx_a]
            blk_b = blocks_b[idx_b]
            loc = ChangeLocation(
                document_id=doc_id,
                block_id=blk_b.id,
                paragraph_id=blk_b.id if isinstance(blk_b, ParagraphBlock) else None,
                text_anchor=getattr(blk_a, "full_text", "")[:60],
                semantic_hash=ChangeLocation.compute_semantic_hash(getattr(blk_a, "full_text", blk_a.id)),
            )
            changes.append(
                ChangeObject(
                    change_id=generate_change_id("MOV"),
                    document_id=doc_id,
                    version_before=v_before,
                    version_after=v_after,
                    change_type=ChangeType.MOVE_PARAGRAPH,
                    target_element=blk_b.id,
                    location=loc,
                    before_content=f"Vị trí cũ: Index {idx_a}",
                    after_content=f"Vị trí mới: Index {idx_b}",
                    reason=f"Di chuyển khối nội dung để cải thiện mạch lạc văn bản (độ tương đồng {round(sim*100)}%).",
                    agent_id=agent_id,
                    agent_action="reorder_content",
                    status=ChangeStatus.PROPOSED,
                    confidence=confidence,
                    evidence=evidence or "Evidence: NOT AVAILABLE",
                )
            )

        # Process Paired/Modified Blocks
        for idx_a, idx_b in matched_a.items():
            blk_a = blocks_a[idx_a]
            blk_b = blocks_b[idx_b]
            block_changes = cls._diff_single_block(
                blk_a, blk_b, doc_id, v_before, v_after, agent_id, default_reason, confidence, evidence
            )
            changes.extend(block_changes)

        # Process Added Blocks (in B not matched in A)
        unmatched_b = [i for i in range(len(blocks_b)) if i not in matched_b]
        for idx_b in unmatched_b:
            blk_b = blocks_b[idx_b]
            changes.append(cls._create_block_added_change(
                blk_b, doc_id, v_before, v_after, agent_id, default_reason, confidence, evidence
            ))

        # Process Deleted Blocks (in A not matched in B)
        unmatched_a = [i for i in range(len(blocks_a)) if i not in matched_a]
        for idx_a in unmatched_a:
            blk_a = blocks_a[idx_a]
            changes.append(cls._create_block_deleted_change(
                blk_a, doc_id, v_before, v_after, agent_id, default_reason, confidence, evidence
            ))

        return changes

    @classmethod
    def _diff_sections(
        cls,
        doc_before: DocumentNode,
        doc_after: DocumentNode,
        doc_id: str,
        v_before: int,
        v_after: int,
        agent_id: str,
        reason: str,
        confidence: float,
        evidence: Optional[str],
    ) -> List[ChangeObject]:
        changes = []
        max_secs = max(len(doc_before.sections), len(doc_after.sections))

        for s_idx in range(max_secs):
            sa = doc_before.sections[s_idx] if s_idx < len(doc_before.sections) else None
            sb = doc_after.sections[s_idx] if s_idx < len(doc_after.sections) else None

            if sa and sb:
                # Header diff
                if sa.header_text != sb.header_text:
                    changes.append(
                        ChangeObject(
                            change_id=generate_change_id("HDR"),
                            document_id=doc_id,
                            version_before=v_before,
                            version_after=v_after,
                            change_type=ChangeType.MODIFY_HEADER,
                            target_element=sb.id,
                            location=ChangeLocation(document_id=doc_id, section_id=sb.id),
                            before_content=sa.header_text,
                            after_content=sb.header_text,
                            reason="Cập nhật tiêu đề đầu trang (Header) cho phân đoạn.",
                            agent_id=agent_id,
                            agent_action="modify_header",
                            confidence=confidence,
                            evidence=evidence or "Evidence: NOT AVAILABLE",
                        )
                    )
                # Footer diff
                if sa.footer_text != sb.footer_text:
                    changes.append(
                        ChangeObject(
                            change_id=generate_change_id("FTR"),
                            document_id=doc_id,
                            version_before=v_before,
                            version_after=v_after,
                            change_type=ChangeType.MODIFY_FOOTER,
                            target_element=sb.id,
                            location=ChangeLocation(document_id=doc_id, section_id=sb.id),
                            before_content=sa.footer_text,
                            after_content=sb.footer_text,
                            reason="Cập nhật chân trang (Footer) cho phân đoạn.",
                            agent_id=agent_id,
                            agent_action="modify_footer",
                            confidence=confidence,
                            evidence=evidence or "Evidence: NOT AVAILABLE",
                        )
                    )
                # Margins diff
                if sa.properties != sb.properties:
                    changes.append(
                        ChangeObject(
                            change_id=generate_change_id("SEC"),
                            document_id=doc_id,
                            version_before=v_before,
                            version_after=v_after,
                            change_type=ChangeType.CHANGE_SECTION,
                            target_element=sb.id,
                            location=ChangeLocation(document_id=doc_id, section_id=sb.id),
                            before_content=sa.properties.model_dump(),
                            after_content=sb.properties.model_dump(),
                            reason="Điều chỉnh thông số căn lề và bố cục trang.",
                            agent_id=agent_id,
                            agent_action="modify_section_properties",
                            affected_structure={"section_properties": sb.properties.model_dump()},
                            confidence=confidence,
                            evidence=evidence or "Evidence: NOT AVAILABLE",
                        )
                    )
            elif sb and not sa:
                changes.append(
                    ChangeObject(
                        change_id=generate_change_id("SEC"),
                        document_id=doc_id,
                        version_before=v_before,
                        version_after=v_after,
                        change_type=ChangeType.CHANGE_SECTION,
                        target_element=sb.id,
                        location=ChangeLocation(document_id=doc_id, section_id=sb.id),
                        before_content=None,
                        after_content=sb.properties.model_dump(),
                        reason="Thêm phân đoạn (Section) mới vào tài liệu.",
                        agent_id=agent_id,
                        agent_action="add_section",
                        confidence=confidence,
                        evidence=evidence or "Evidence: NOT AVAILABLE",
                    )
                )

        return changes

    @classmethod
    def _match_blocks(
        cls,
        blocks_a: List[BaseBlockNode],
        blocks_b: List[BaseBlockNode],
    ) -> Tuple[Dict[int, int], Dict[int, int], List[Tuple[int, int, float]]]:
        """
        Pairs blocks between before and after versions using stable IDs and semantic similarity.
        Detects block reordering / moving without misclassifying as delete+insert.
        """
        matched_a: Dict[int, int] = {}
        matched_b: Dict[int, int] = {}
        moves: List[Tuple[int, int, float]] = []

        # 1. Exact ID matching
        ids_b = {blk.id: idx for idx, blk in enumerate(blocks_b)}
        for idx_a, blk_a in enumerate(blocks_a):
            if blk_a.id in ids_b:
                idx_b = ids_b[blk_a.id]
                matched_a[idx_a] = idx_b
                matched_b[idx_b] = idx_a

        # 2. Semantic matching for unmatched blocks (detect edits with new IDs or moves)
        unmatched_a_indices = [i for i in range(len(blocks_a)) if i not in matched_a]
        unmatched_b_indices = [i for i in range(len(blocks_b)) if i not in matched_b]

        for idx_a in unmatched_a_indices:
            blk_a = blocks_a[idx_a]
            txt_a = getattr(blk_a, "full_text", getattr(blk_a, "text", ""))
            if not txt_a:
                continue

            best_match = None
            best_sim = 0.0

            for idx_b in unmatched_b_indices:
                if idx_b in matched_b:
                    continue
                blk_b = blocks_b[idx_b]
                txt_b = getattr(blk_b, "full_text", getattr(blk_b, "text", ""))
                if not txt_b:
                    continue

                if type(blk_a) is type(blk_b):
                    sim = difflib.SequenceMatcher(None, txt_a, txt_b).ratio()
                    if sim > best_sim and sim >= 0.60:
                        best_sim = sim
                        best_match = idx_b

            if best_match is not None:
                matched_a[idx_a] = best_match
                matched_b[best_match] = idx_a
                # If similarity is very high but position shifted significantly, record as move
                if best_sim >= cls.SIMILARITY_MOVE_THRESHOLD and abs(idx_a - best_match) > 1:
                    moves.append((idx_a, best_match, best_sim))

        return matched_a, matched_b, moves

    @classmethod
    def _diff_single_block(
        cls,
        blk_a: BaseBlockNode,
        blk_b: BaseBlockNode,
        doc_id: str,
        v_before: int,
        v_after: int,
        agent_id: str,
        reason: str,
        confidence: float,
        evidence: Optional[str],
    ) -> List[ChangeObject]:
        changes: List[ChangeObject] = []

        # 1. Heading Diff
        if isinstance(blk_a, HeadingBlock) and isinstance(blk_b, HeadingBlock):
            txt_a = blk_a.full_text
            txt_b = blk_b.full_text
            if blk_a.level != blk_b.level:
                changes.append(
                    ChangeObject(
                        change_id=generate_change_id("HDG"),
                        document_id=doc_id,
                        version_before=v_before,
                        version_after=v_after,
                        change_type=ChangeType.CHANGE_HEADING_LEVEL,
                        target_element=blk_b.id,
                        location=ChangeLocation(document_id=doc_id, block_id=blk_b.id, text_anchor=txt_a[:60]),
                        before_content=f"Heading {blk_a.level} (Cấp {blk_a.level})",
                        after_content=f"Heading {blk_b.level} (Cấp {blk_b.level})",
                        reason=f"Điều chỉnh cấp độ tiêu đề từ Heading {blk_a.level} sang Heading {blk_b.level}.",
                        agent_id=agent_id,
                        agent_action="change_heading_level",
                        affected_structure={"old_level": blk_a.level, "new_level": blk_b.level},
                        confidence=confidence,
                        evidence=evidence or "Evidence: NOT AVAILABLE",
                    )
                )
            if txt_a != txt_b:
                changes.append(
                    ChangeObject(
                        change_id=generate_change_id("HDG"),
                        document_id=doc_id,
                        version_before=v_before,
                        version_after=v_after,
                        change_type=ChangeType.MODIFY_HEADING_TEXT,
                        target_element=blk_b.id,
                        location=ChangeLocation(document_id=doc_id, block_id=blk_b.id, text_anchor=txt_a[:60]),
                        before_content=txt_a,
                        after_content=txt_b,
                        reason="Cập nhật câu chữ tiêu đề rõ ràng, mạch lạc hơn.",
                        agent_id=agent_id,
                        agent_action="modify_heading_text",
                        confidence=confidence,
                        evidence=evidence or "Evidence: NOT AVAILABLE",
                    )
                )
            # Paragraph properties diff on heading
            changes.extend(cls._diff_paragraph_properties(blk_a, blk_b, doc_id, v_before, v_after, agent_id, confidence, evidence))
            return changes

        # 2. Paragraph Diff (Text & Formatting)
        if isinstance(blk_a, ParagraphBlock) and isinstance(blk_b, ParagraphBlock):
            txt_a = blk_a.full_text
            txt_b = blk_b.full_text

            if txt_a != txt_b:
                text_changes = cls._diff_paragraph_text(
                    blk_a, blk_b, doc_id, v_before, v_after, agent_id, reason, confidence, evidence
                )
                changes.extend(text_changes)

            # Paragraph formatting diff (alignment, spacing, indents)
            changes.extend(cls._diff_paragraph_properties(blk_a, blk_b, doc_id, v_before, v_after, agent_id, confidence, evidence))
            return changes

        # 3. Table Diff
        if isinstance(blk_a, TableBlock) and isinstance(blk_b, TableBlock):
            changes.extend(cls._diff_tables(blk_a, blk_b, doc_id, v_before, v_after, agent_id, reason, confidence, evidence))
            return changes

        return changes

    @classmethod
    def _diff_paragraph_text(
        cls,
        blk_a: ParagraphBlock,
        blk_b: ParagraphBlock,
        doc_id: str,
        v_before: int,
        v_after: int,
        agent_id: str,
        reason: str,
        confidence: float,
        evidence: Optional[str],
    ) -> List[ChangeObject]:
        changes = []
        txt_a = blk_a.full_text
        txt_b = blk_b.full_text

        matcher = difflib.SequenceMatcher(None, txt_a, txt_b)
        opcodes = matcher.get_opcodes()

        # If many fragmented edits, group into a clean paragraph replacement change
        if len(opcodes) > 4 or (len(txt_a) > 0 and len(txt_b) > 0 and matcher.ratio() < 0.4):
            changes.append(
                ChangeObject(
                    change_id=generate_change_id("TXT"),
                    document_id=doc_id,
                    version_before=v_before,
                    version_after=v_after,
                    change_type=ChangeType.MODIFY_PARAGRAPH,
                    target_element=blk_b.id,
                    location=ChangeLocation(
                        document_id=doc_id,
                        block_id=blk_b.id,
                        paragraph_id=blk_b.id,
                        start_offset=0,
                        end_offset=len(txt_b),
                        text_anchor=txt_a[:60],
                        semantic_hash=ChangeLocation.compute_semantic_hash(txt_a),
                    ),
                    before_content=txt_a,
                    after_content=txt_b,
                    reason=reason,
                    agent_id=agent_id,
                    agent_action="rewrite_paragraph",
                    confidence=confidence,
                    evidence=evidence or "Evidence: NOT AVAILABLE",
                )
            )
            return changes

        # Fine-grained word/character edits
        for tag, i1, i2, j1, j2 in opcodes:
            if tag == "equal":
                continue

            sub_before = txt_a[i1:i2]
            sub_after = txt_b[j1:j2]
            chg_type = ChangeType.REPLACE_TEXT
            if tag == "insert":
                chg_type = ChangeType.INSERT_TEXT
            elif tag == "delete":
                chg_type = ChangeType.DELETE_TEXT

            changes.append(
                ChangeObject(
                    change_id=generate_change_id("TXT"),
                    document_id=doc_id,
                    version_before=v_before,
                    version_after=v_after,
                    change_type=chg_type,
                    target_element=blk_b.id,
                    location=ChangeLocation(
                        document_id=doc_id,
                        block_id=blk_b.id,
                        paragraph_id=blk_b.id,
                        start_offset=j1,
                        end_offset=j2,
                        text_anchor=txt_a[max(0, i1 - 20) : min(len(txt_a), i2 + 20)],
                        semantic_hash=ChangeLocation.compute_semantic_hash(sub_before or sub_after),
                    ),
                    before_content=sub_before if sub_before else None,
                    after_content=sub_after if sub_after else None,
                    reason=reason,
                    agent_id=agent_id,
                    agent_action="edit_text",
                    confidence=confidence,
                    evidence=evidence or "Evidence: NOT AVAILABLE",
                )
            )

        return changes

    @classmethod
    def _diff_paragraph_properties(
        cls,
        blk_a: ParagraphBlock,
        blk_b: ParagraphBlock,
        doc_id: str,
        v_before: int,
        v_after: int,
        agent_id: str,
        confidence: float,
        evidence: Optional[str],
    ) -> List[ChangeObject]:
        changes = []
        fmt_diff: Dict[str, Any] = {}

        if blk_a.alignment != blk_b.alignment:
            fmt_diff["alignment"] = {"before": blk_a.alignment, "after": blk_b.alignment}
        if blk_a.line_spacing != blk_b.line_spacing:
            fmt_diff["line_spacing"] = {"before": blk_a.line_spacing, "after": blk_b.line_spacing}
        if blk_a.space_before_pt != blk_b.space_before_pt:
            fmt_diff["space_before_pt"] = {"before": blk_a.space_before_pt, "after": blk_b.space_before_pt}
        if blk_a.space_after_pt != blk_b.space_after_pt:
            fmt_diff["space_after_pt"] = {"before": blk_a.space_after_pt, "after": blk_b.space_after_pt}
        if blk_a.first_line_indent_cm != blk_b.first_line_indent_cm:
            fmt_diff["first_line_indent_cm"] = {"before": blk_a.first_line_indent_cm, "after": blk_b.first_line_indent_cm}

        if fmt_diff:
            changes.append(
                ChangeObject(
                    change_id=generate_change_id("FMT"),
                    document_id=doc_id,
                    version_before=v_before,
                    version_after=v_after,
                    change_type=ChangeType.FORMAT_PARAGRAPH,
                    target_element=blk_b.id,
                    location=ChangeLocation(document_id=doc_id, block_id=blk_b.id, paragraph_id=blk_b.id),
                    before_content={k: v["before"] for k, v in fmt_diff.items()},
                    after_content={k: v["after"] for k, v in fmt_diff.items()},
                    reason="Căn chỉnh khoảng cách đoạn văn theo tiêu chuẩn học thuật.",
                    agent_id=agent_id,
                    agent_action="format_paragraph",
                    affected_formatting=fmt_diff,
                    confidence=confidence,
                    evidence=evidence or "Evidence: NOT AVAILABLE",
                )
            )

        return changes

    @classmethod
    def _diff_tables(
        cls,
        tbl_a: TableBlock,
        tbl_b: TableBlock,
        doc_id: str,
        v_before: int,
        v_after: int,
        agent_id: str,
        reason: str,
        confidence: float,
        evidence: Optional[str],
    ) -> List[ChangeObject]:
        changes = []

        # Row counts
        if tbl_a.rows != tbl_b.rows:
            chg_t = ChangeType.ADD_ROW if tbl_b.rows > tbl_a.rows else ChangeType.DELETE_ROW
            changes.append(
                ChangeObject(
                    change_id=generate_change_id("TBL"),
                    document_id=doc_id,
                    version_before=v_before,
                    version_after=v_after,
                    change_type=chg_t,
                    target_element=tbl_b.id,
                    location=ChangeLocation(document_id=doc_id, block_id=tbl_b.id, table_id=tbl_b.id),
                    before_content=f"{tbl_a.rows} hàng",
                    after_content=f"{tbl_b.rows} hàng",
                    reason=f"Thay đổi số lượng hàng trong bảng ({tbl_a.rows} -> {tbl_b.rows}).",
                    agent_id=agent_id,
                    agent_action="modify_table_rows",
                    confidence=confidence,
                    evidence=evidence or "Evidence: NOT AVAILABLE",
                )
            )

        # Cell Content & Formatting Diffs
        min_rows = min(len(tbl_a.cells), len(tbl_b.cells))
        for r_idx in range(min_rows):
            row_a = tbl_a.cells[r_idx]
            row_b = tbl_b.cells[r_idx]
            min_cols = min(len(row_a), len(row_b))
            for c_idx in range(min_cols):
                c_a = row_a[c_idx]
                c_b = row_b[c_idx]

                if c_a.text != c_b.text:
                    changes.append(
                        ChangeObject(
                            change_id=generate_change_id("CEL"),
                            document_id=doc_id,
                            version_before=v_before,
                            version_after=v_after,
                            change_type=ChangeType.MODIFY_CELL_CONTENT,
                            target_element=c_b.id or tbl_b.id,
                            location=ChangeLocation(
                                document_id=doc_id,
                                block_id=tbl_b.id,
                                table_id=tbl_b.id,
                                row_idx=r_idx,
                                col_idx=c_idx,
                                cell_id=c_b.id,
                            ),
                            before_content=c_a.text,
                            after_content=c_b.text,
                            reason=f"Cập nhật nội dung ô hàng {r_idx+1}, cột {c_idx+1}.",
                            agent_id=agent_id,
                            agent_action="edit_table_cell",
                            confidence=confidence,
                            evidence=evidence or "Evidence: NOT AVAILABLE",
                        )
                    )

                if c_a.bg_color_hex != c_b.bg_color_hex:
                    changes.append(
                        ChangeObject(
                            change_id=generate_change_id("FMT"),
                            document_id=doc_id,
                            version_before=v_before,
                            version_after=v_after,
                            change_type=ChangeType.MODIFY_CELL_FORMATTING,
                            target_element=c_b.id or tbl_b.id,
                            location=ChangeLocation(
                                document_id=doc_id,
                                block_id=tbl_b.id,
                                table_id=tbl_b.id,
                                row_idx=r_idx,
                                col_idx=c_idx,
                                cell_id=c_b.id,
                            ),
                            before_content=c_a.bg_color_hex,
                            after_content=c_b.bg_color_hex,
                            reason=f"Tô màu nền ô hàng {r_idx+1}, cột {c_idx+1}.",
                            agent_id=agent_id,
                            agent_action="format_cell_background",
                            affected_formatting={"bg_color_hex": c_b.bg_color_hex},
                            confidence=confidence,
                            evidence=evidence or "Evidence: NOT AVAILABLE",
                        )
                    )

        return changes

    @classmethod
    def _create_block_added_change(
        cls,
        blk: BaseBlockNode,
        doc_id: str,
        v_before: int,
        v_after: int,
        agent_id: str,
        reason: str,
        confidence: float,
        evidence: Optional[str],
    ) -> ChangeObject:
        chg_type = ChangeType.ADD_PARAGRAPH
        if isinstance(blk, HeadingBlock):
            chg_type = ChangeType.ADD_HEADING
        elif isinstance(blk, TableBlock):
            chg_type = ChangeType.ADD_TABLE
        elif isinstance(blk, ImageBlock):
            chg_type = ChangeType.ADD_IMAGE

        txt = getattr(blk, "full_text", getattr(blk, "text", str(blk.id)))

        return ChangeObject(
            change_id=generate_change_id("ADD"),
            document_id=doc_id,
            version_before=v_before,
            version_after=v_after,
            change_type=chg_type,
            target_element=blk.id,
            location=ChangeLocation(document_id=doc_id, block_id=blk.id, text_anchor=txt[:60]),
            before_content=None,
            after_content=txt,
            reason=f"Chèn thêm {chg_type.value.lower().replace('_', ' ')} vào tài liệu.",
            agent_id=agent_id,
            agent_action="insert_element",
            confidence=confidence,
            evidence=evidence or "Evidence: NOT AVAILABLE",
        )

    @classmethod
    def _create_block_deleted_change(
        cls,
        blk: BaseBlockNode,
        doc_id: str,
        v_before: int,
        v_after: int,
        agent_id: str,
        reason: str,
        confidence: float,
        evidence: Optional[str],
    ) -> ChangeObject:
        chg_type = ChangeType.DELETE_PARAGRAPH
        if isinstance(blk, HeadingBlock):
            chg_type = ChangeType.REMOVE_HEADING
        elif isinstance(blk, TableBlock):
            chg_type = ChangeType.DELETE_TABLE
        elif isinstance(blk, ImageBlock):
            chg_type = ChangeType.DELETE_IMAGE

        txt = getattr(blk, "full_text", getattr(blk, "text", str(blk.id)))

        return ChangeObject(
            change_id=generate_change_id("DEL"),
            document_id=doc_id,
            version_before=v_before,
            version_after=v_after,
            change_type=chg_type,
            target_element=blk.id,
            location=ChangeLocation(document_id=doc_id, block_id=blk.id, text_anchor=txt[:60]),
            before_content=txt,
            after_content=None,
            reason=f"Xóa {chg_type.value.lower().replace('_', ' ')} dư thừa khỏi tài liệu.",
            agent_id=agent_id,
            agent_action="delete_element",
            confidence=confidence,
            evidence=evidence or "Evidence: NOT AVAILABLE",
        )
