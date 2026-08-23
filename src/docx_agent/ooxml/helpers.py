"""
OOXML helper utilities for XML manipulation, namespaces, and node cloning.
"""

from copy import deepcopy
from typing import Optional, List, Tuple, Any
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml


def set_run_text_preserving_spaces(run_element, text: str) -> None:
    """
    Sets text on a w:r element, ensuring w:t has xml:space="preserve"
    if leading/trailing whitespace exists.
    """
    t_elems = run_element.xpath("./w:t")
    if t_elems:
        t = t_elems[0]
        t.text = text
        if text.startswith(" ") or text.endswith(" ") or "  " in text:
            t.set(qn("xml:space"), "preserve")
        # Remove extra w:t if any
        for extra in t_elems[1:]:
            run_element.remove(extra)
    else:
        t = OxmlElement("w:t")
        t.text = text
        if text.startswith(" ") or text.endswith(" ") or "  " in text:
            t.set(qn("xml:space"), "preserve")
        run_element.append(t)


def clone_run_element(run_element) -> Any:
    """Deep-clones a w:r XML element."""
    return deepcopy(run_element)
