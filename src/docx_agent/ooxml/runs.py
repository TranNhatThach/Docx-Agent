"""
Run Surgery Engine for format-preserving text replacements across run boundaries.
"""

from typing import List, Tuple, Optional, Any
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from docx_agent.ooxml.helpers import clone_run_element, set_run_text_preserving_spaces
from docx_agent.utils.unicode import normalize_unicode


def get_paragraph_runs(p: Paragraph) -> List[Any]:
    """Returns all w:r elements in paragraph order (including inside hyperlinks)."""
    return p._p.xpath(".//w:r")


def get_run_text(r_elem) -> str:
    """Extracts text from all w:t elements in a run."""
    t_elems = r_elem.xpath("./w:t")
    return "".join(t.text or "" for t in t_elems)


def replace_in_paragraph_preserving_formatting(
    p: Paragraph,
    target: str,
    replacement: str,
    count: Optional[int] = None,
) -> int:
    """
    Replaces occurrences of `target` with `replacement` in a paragraph
    while strictly preserving run boundaries and formatting properties (bold, italic, colors, fonts, etc.).
    
    Returns the number of replacements performed.
    """
    if not target:
        return 0

    replacements_done = 0

    while True:
        if count is not None and replacements_done >= count:
            break

        # Collect current runs and text
        r_elems = get_paragraph_runs(p)
        if not r_elems:
            break

        # Build run segments and character map
        char_map: List[Tuple[Any, int, str]] = []  # (r_elem, offset_in_run, full_run_text)
        full_text_chars: List[str] = []

        for r_elem in r_elems:
            r_text = get_run_text(r_elem)
            for offset, char in enumerate(r_text):
                char_map.append((r_elem, offset, r_text))
                full_text_chars.append(char)

        full_text = "".join(full_text_chars)
        match_idx = full_text.find(target)
        if match_idx == -1:
            break

        match_start = match_idx
        match_end = match_idx + len(target)

        start_elem, start_offset, start_text = char_map[match_start]
        end_elem, end_offset, end_text = char_map[match_end - 1]

        if start_elem is end_elem:
            # Match is entirely within a single run
            prefix = start_text[:start_offset]
            suffix = start_text[start_offset + len(target) :]

            # Insert clone for suffix if non-empty
            if suffix:
                suffix_run = clone_run_element(start_elem)
                set_run_text_preserving_spaces(suffix_run, suffix)
                start_elem.addnext(suffix_run)

            # Insert clone for replacement if prefix exists, or edit in-place
            if prefix:
                set_run_text_preserving_spaces(start_elem, prefix)
                repl_run = clone_run_element(start_elem)
                set_run_text_preserving_spaces(repl_run, replacement)
                start_elem.addnext(repl_run)
            else:
                set_run_text_preserving_spaces(start_elem, replacement)

        else:
            # Match spans across multiple runs
            # 1. Start run: Keep prefix, append replacement
            prefix = start_text[:start_offset]
            if prefix:
                set_run_text_preserving_spaces(start_elem, prefix)
                repl_run = clone_run_element(start_elem)
                set_run_text_preserving_spaces(repl_run, replacement)
                start_elem.addnext(repl_run)
            else:
                set_run_text_preserving_spaces(start_elem, replacement)

            # 2. Intermediate runs: Remove text
            in_between = False
            for r_elem in r_elems:
                if r_elem is start_elem:
                    in_between = True
                    continue
                if r_elem is end_elem:
                    break
                if in_between:
                    set_run_text_preserving_spaces(r_elem, "")

            # 3. End run: Keep suffix
            suffix = end_text[end_offset + 1 :]
            set_run_text_preserving_spaces(end_elem, suffix)

        replacements_done += 1

    return replacements_done
