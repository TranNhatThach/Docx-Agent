"""
Format verification engine: Checks document typography conformance against target specifications.
"""

from typing import Dict, Any, List, Optional, Union
from pathlib import Path
import docx
from pydantic import BaseModel, Field
from docx_agent.utils.paths import resolve_safe_path


class FormatVerificationReport(BaseModel):
    passed: bool
    paragraphs_checked: int = 0
    font_correct: int = 0
    size_correct: int = 0
    line_spacing_correct: int = 0
    alignment_correct: int = 0
    failed: int = 0
    failures: List[Dict[str, Any]] = Field(default_factory=list)


class FormatChecker:
    """
    Validates paragraph fonts, sizes, alignments, and spacings against expected criteria.
    """

    @staticmethod
    def verify_document_formatting(
        doc_or_path: Union[str, docx.Document],
        expected_font: Optional[str] = None,
        expected_size_pt: Optional[float] = None,
        expected_line_spacing: Optional[float] = None,
        expected_alignment: Optional[str] = None,
        ignore_headings: bool = True,
    ) -> FormatVerificationReport:
        """
        Scans paragraphs and verifies adherence to expected typography properties.
        """
        if isinstance(doc_or_path, (str, Path)):
            doc = docx.Document(str(resolve_safe_path(doc_or_path)))
        else:
            doc = doc_or_path

        total_checked = 0
        font_ok = 0
        size_ok = 0
        spacing_ok = 0
        align_ok = 0
        failures = []

        for idx, p in enumerate(doc.paragraphs):
            style_name = p.style.name if p.style else ""
            if ignore_headings and ("heading" in style_name.lower() or style_name in ["Title", "Subtitle"]):
                continue

            if not p.text.strip():
                continue

            pid = f"p_{idx + 1:04d}"
            total_checked += 1
            p_has_failure = False

            # 1. Font name check (inspect runs or style)
            if expected_font:
                matched_font = False
                for r in p.runs:
                    if r.font and r.font.name:
                        if r.font.name.lower() == expected_font.lower():
                            matched_font = True
                            break
                if not matched_font and p.style and p.style.font and p.style.font.name:
                    if p.style.font.name.lower() == expected_font.lower():
                        matched_font = True
                
                # Check document default or Normal style if run didn't override
                if not matched_font and doc.styles["Normal"].font and doc.styles["Normal"].font.name:
                    if doc.styles["Normal"].font.name.lower() == expected_font.lower():
                        matched_font = True

                if matched_font:
                    font_ok += 1
                else:
                    p_has_failure = True
                    failures.append({
                        "element": pid,
                        "property": "font_name",
                        "expected": expected_font,
                        "actual": p.runs[0].font.name if (p.runs and p.runs[0].font and p.runs[0].font.name) else "Unknown",
                    })

            # 2. Font size check
            if expected_size_pt is not None:
                matched_size = False
                for r in p.runs:
                    if r.font and r.font.size and abs(r.font.size.pt - expected_size_pt) < 0.1:
                        matched_size = True
                        break
                if not matched_size and p.style and p.style.font and p.style.font.size:
                    if abs(p.style.font.size.pt - expected_size_pt) < 0.1:
                        matched_size = True
                if not matched_size and doc.styles["Normal"].font and doc.styles["Normal"].font.size:
                    if abs(doc.styles["Normal"].font.size.pt - expected_size_pt) < 0.1:
                        matched_size = True

                if matched_size:
                    size_ok += 1
                else:
                    p_has_failure = True
                    failures.append({
                        "element": pid,
                        "property": "font_size_pt",
                        "expected": expected_size_pt,
                        "actual": p.runs[0].font.size.pt if (p.runs and p.runs[0].font and p.runs[0].font.size) else None,
                    })

            # 3. Line spacing check
            if expected_line_spacing is not None:
                ls = p.paragraph_format.line_spacing
                if ls is None and p.style and hasattr(p.style, "paragraph_format"):
                    ls = p.style.paragraph_format.line_spacing
                if ls is None and "Normal" in doc.styles and hasattr(doc.styles["Normal"], "paragraph_format"):
                    ls = doc.styles["Normal"].paragraph_format.line_spacing

                if ls is not None and abs(ls - expected_line_spacing) < 0.05:
                    spacing_ok += 1
                else:
                    p_has_failure = True
                    failures.append({
                        "element": pid,
                        "property": "line_spacing",
                        "expected": expected_line_spacing,
                        "actual": ls,
                    })

            # 4. Alignment check
            if expected_alignment is not None:
                align_str = str(p.alignment).split(".")[-1].lower() if p.alignment is not None else None
                if align_str is None and p.style and hasattr(p.style, "paragraph_format") and p.style.paragraph_format.alignment is not None:
                    align_str = str(p.style.paragraph_format.alignment).split(".")[-1].lower()
                if align_str is None and "Normal" in doc.styles and hasattr(doc.styles["Normal"], "paragraph_format") and doc.styles["Normal"].paragraph_format.alignment is not None:
                    align_str = str(doc.styles["Normal"].paragraph_format.alignment).split(".")[-1].lower()

                if align_str and align_str == expected_alignment.lower():
                    align_ok += 1
                else:
                    p_has_failure = True
                    failures.append({
                        "element": pid,
                        "property": "alignment",
                        "expected": expected_alignment,
                        "actual": align_str or "left",
                    })

        return FormatVerificationReport(
            passed=len(failures) == 0,
            paragraphs_checked=total_checked,
            font_correct=font_ok,
            size_correct=size_ok,
            line_spacing_correct=spacing_ok,
            alignment_correct=align_ok,
            failed=len(failures),
            failures=failures,
        )
