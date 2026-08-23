"""
Visual Layout Verification Engine: Detects visual layout anomalies,
table/image boundary overflows, orphan headings, and typography inconsistencies.
"""

from typing import List, Dict, Any, Optional
from pydantic import BaseModel, Field

from docx_agent.canonical.model import (
    DocumentNode,
    HeadingBlock,
    ParagraphBlock,
    TableBlock,
    ImageBlock,
    DiagramBlock,
    SectionNode,
)


class VisualAnomaly(BaseModel):
    element_id: str
    anomaly_type: str  # "IMAGE_OVERFLOW", "TABLE_OVERFLOW", "HEADING_HIERARCHY", "EMPTY_TRAILING_PAGE", "EXCESSIVE_WHITESPACE"
    severity: str = "WARNING"  # "ERROR", "WARNING", "INFO"
    message: str
    suggested_fix: str


class VisualVerificationReport(BaseModel):
    passed: bool
    total_blocks_checked: int = 0
    anomalies: List[VisualAnomaly] = Field(default_factory=list)
    layout_score: float = 1.0


class VisualLayoutVerifier:
    """
    Analyzes document layout boundaries and visual constraints.
    """

    @staticmethod
    def verify_document_layout(doc: DocumentNode) -> VisualVerificationReport:
        anomalies: List[VisualAnomaly] = []
        total_blocks = 0
        last_heading_level = 0

        for sec in doc.sections:
            printable_width_cm = (
                sec.properties.page_width_cm
                - sec.properties.margin_left_cm
                - sec.properties.margin_right_cm
            )

            for blk in sec.blocks:
                total_blocks += 1

                # 1. Heading Hierarchy Check
                if isinstance(blk, HeadingBlock):
                    if blk.level > last_heading_level + 1 and last_heading_level > 0:
                        anomalies.append(
                            VisualAnomaly(
                                element_id=blk.id,
                                anomaly_type="HEADING_HIERARCHY",
                                severity="WARNING",
                                message=f"Heading level skipped: H{last_heading_level} followed immediately by H{blk.level}.",
                                suggested_fix=f"Change heading level to H{last_heading_level + 1} for consistent document hierarchy.",
                            )
                        )
                    last_heading_level = blk.level

                # 2. Image Overflow Check
                elif isinstance(blk, ImageBlock):
                    if blk.width_cm and blk.width_cm > printable_width_cm:
                        anomalies.append(
                            VisualAnomaly(
                                element_id=blk.id,
                                anomaly_type="IMAGE_OVERFLOW",
                                severity="ERROR",
                                message=f"Image width ({blk.width_cm:.1f}cm) exceeds printable page width ({printable_width_cm:.1f}cm).",
                                suggested_fix=f"Resize image width to <= {printable_width_cm:.1f}cm.",
                            )
                        )

                # 3. Diagram Width Check
                elif isinstance(blk, DiagramBlock):
                    if blk.width_cm and blk.width_cm > printable_width_cm:
                        anomalies.append(
                            VisualAnomaly(
                                element_id=blk.id,
                                anomaly_type="DIAGRAM_OVERFLOW",
                                severity="WARNING",
                                message=f"Diagram width ({blk.width_cm:.1f}cm) exceeds printable page width ({printable_width_cm:.1f}cm).",
                                suggested_fix=f"Adjust diagram width to <= {printable_width_cm:.1f}cm.",
                            )
                        )

                # 4. Table Column Overflow Check
                elif isinstance(blk, TableBlock):
                    if blk.columns > 10:
                        anomalies.append(
                            VisualAnomaly(
                                element_id=blk.id,
                                anomaly_type="TABLE_OVERFLOW",
                                severity="WARNING",
                                message=f"Table contains {blk.columns} columns which may clip on portrait A4 page.",
                                suggested_fix="Switch section to landscape orientation or reduce table columns.",
                            )
                        )

                # 5. Excessive Spacing Check
                elif isinstance(blk, ParagraphBlock):
                    if blk.space_before_pt and blk.space_before_pt > 72.0:
                        anomalies.append(
                            VisualAnomaly(
                                element_id=blk.id,
                                anomaly_type="EXCESSIVE_WHITESPACE",
                                severity="WARNING",
                                message=f"Paragraph space before ({blk.space_before_pt}pt) creates excessive gap.",
                                suggested_fix="Reduce paragraph space before to <= 24pt.",
                            )
                        )

        # Check for trailing empty blocks
        if sec.blocks and isinstance(sec.blocks[-1], ParagraphBlock) and not sec.blocks[-1].full_text.strip():
            anomalies.append(
                VisualAnomaly(
                    element_id=sec.blocks[-1].id,
                    anomaly_type="EMPTY_TRAILING_PAGE",
                    severity="INFO",
                    message="Document ends with an empty paragraph block.",
                    suggested_fix="Remove trailing empty paragraph to prevent unwanted blank pages.",
                )
            )

        errors_count = sum(1 for a in anomalies if a.severity == "ERROR")
        passed = errors_count == 0

        score = max(0.0, 1.0 - (len(anomalies) * 0.1))

        return VisualVerificationReport(
            passed=passed,
            total_blocks_checked=total_blocks,
            anomalies=anomalies,
            layout_score=round(score, 2),
        )
