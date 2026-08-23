"""
NumberingResolver: High-Fidelity WordprocessingML Numbering Engine.
Parses numbering.xml (abstractNum, num, lvl, numFmt, lvlText, start) and generates deterministic list prefixes.
"""

from typing import Dict, Any, Optional, List
from xml.etree import ElementTree as ET


class NumberingLevel:
    def __init__(
        self,
        ilvl: int,
        num_fmt: str = "decimal",
        lvl_text: str = "%1.",
        start: int = 1,
    ):
        self.ilvl = ilvl
        self.num_fmt = num_fmt  # "decimal", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman", "bullet", "none"
        self.lvl_text = lvl_text
        self.start = start


class AbstractNumbering:
    def __init__(self, abstract_num_id: int):
        self.abstract_num_id = abstract_num_id
        self.levels: Dict[int, NumberingLevel] = {}


class NumberingResolver:
    """
    Evaluates dynamic numbering counters across document paragraphs.
    """

    def __init__(self, raw_numbering_xml: Optional[str] = None):
        self.abstract_nums: Dict[int, AbstractNumbering] = {}
        self.num_to_abstract: Dict[int, int] = {}
        self.counters: Dict[int, Dict[int, int]] = {}  # num_id -> {ilvl: current_val}
        if raw_numbering_xml:
            self.parse_numbering_xml(raw_numbering_xml)

    def parse_numbering_xml(self, xml_content: str) -> None:
        try:
            root = ET.fromstring(xml_content)
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

            # Parse abstractNum
            for abs_elem in root.findall("w:abstractNum", ns):
                abs_id_str = abs_elem.attrib.get(f"{{{ns['w']}}}abstractNumId") or abs_elem.attrib.get("abstractNumId")
                if not abs_id_str or not abs_id_str.isdigit():
                    continue
                abs_id = int(abs_id_str)
                abs_num = AbstractNumbering(abs_id)

                for lvl_elem in abs_elem.findall("w:lvl", ns):
                    ilvl_str = lvl_elem.attrib.get(f"{{{ns['w']}}}ilvl") or lvl_elem.attrib.get("ilvl")
                    if not ilvl_str or not ilvl_str.isdigit():
                        continue
                    ilvl = int(ilvl_str)

                    num_fmt = "decimal"
                    fmt_elem = lvl_elem.find("w:numFmt", ns)
                    if fmt_elem is not None:
                        num_fmt = fmt_elem.attrib.get(f"{{{ns['w']}}}val", "decimal")

                    lvl_text = f"%{ilvl+1}."
                    txt_elem = lvl_elem.find("w:lvlText", ns)
                    if txt_elem is not None:
                        lvl_text = txt_elem.attrib.get(f"{{{ns['w']}}}val", f"%{ilvl+1}.")

                    start_val = 1
                    start_elem = lvl_elem.find("w:start", ns)
                    if start_elem is not None:
                        val = start_elem.attrib.get(f"{{{ns['w']}}}val", "1")
                        if val.isdigit():
                            start_val = int(val)

                    abs_num.levels[ilvl] = NumberingLevel(
                        ilvl=ilvl,
                        num_fmt=num_fmt,
                        lvl_text=lvl_text,
                        start=start_val,
                    )

                self.abstract_nums[abs_id] = abs_num

            # Parse num mappings
            for num_elem in root.findall("w:num", ns):
                num_id_str = num_elem.attrib.get(f"{{{ns['w']}}}numId") or num_elem.attrib.get("numId")
                if not num_id_str or not num_id_str.isdigit():
                    continue
                num_id = int(num_id_str)

                abs_ref = num_elem.find("w:abstractNumId", ns)
                if abs_ref is not None:
                    ref_val = abs_ref.attrib.get(f"{{{ns['w']}}}val")
                    if ref_val and ref_val.isdigit():
                        self.num_to_abstract[num_id] = int(ref_val)
        except Exception:
            pass

    def get_numbering_label(self, num_id: int, ilvl: int = 0) -> str:
        """
        Advances the numbering counter and returns the resolved display string (e.g. '1.', '1.1', '•').
        """
        if num_id not in self.num_to_abstract:
            # Fallback decimal
            return f"{self._increment_counter(num_id, ilvl)}."

        abs_id = self.num_to_abstract[num_id]
        if abs_id not in self.abstract_nums:
            return f"{self._increment_counter(num_id, ilvl)}."

        abs_num = self.abstract_nums[abs_id]
        lvl = abs_num.levels.get(ilvl)
        if not lvl:
            return f"{self._increment_counter(num_id, ilvl)}."

        curr_val = self._increment_counter(num_id, ilvl, lvl.start)

        if lvl.num_fmt == "bullet":
            return lvl.lvl_text or "•"

        # Construct multilevel label string
        res = lvl.lvl_text
        for i in range(ilvl + 1):
            val_i = self.counters.get(num_id, {}).get(i, 1)
            formatted_val = self._format_number(val_i, abs_num.levels.get(i, lvl).num_fmt)
            res = res.replace(f"%{i+1}", str(formatted_val))

        return res

    def _increment_counter(self, num_id: int, ilvl: int, start: int = 1) -> int:
        if num_id not in self.counters:
            self.counters[num_id] = {}

        # Reset deeper levels
        for deep_lvl in list(self.counters[num_id].keys()):
            if deep_lvl > ilvl:
                self.counters[num_id][deep_lvl] = 0

        curr = self.counters[num_id].get(ilvl, start - 1) + 1
        self.counters[num_id][ilvl] = curr
        return curr

    def _format_number(self, value: int, num_fmt: str) -> str:
        if num_fmt == "decimal":
            return str(value)
        elif num_fmt == "lowerLetter":
            return chr(ord("a") + (value - 1) % 26)
        elif num_fmt == "upperLetter":
            return chr(ord("A") + (value - 1) % 26)
        elif num_fmt == "lowerRoman":
            return self._int_to_roman(value).lower()
        elif num_fmt == "upperRoman":
            return self._int_to_roman(value).upper()
        return str(value)

    @staticmethod
    def _int_to_roman(num: int) -> str:
        val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
        syb = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"]
        roman_num = ""
        i = 0
        while num > 0:
            for _ in range(num // val[i]):
                roman_num += syb[i]
                num -= val[i]
            i += 1
        return roman_num or "I"
