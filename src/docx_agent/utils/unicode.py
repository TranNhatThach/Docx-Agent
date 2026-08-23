"""
Unicode and Vietnamese language handling utilities.
Ensures perfect normalization (NFC) and preservation of Vietnamese diacritics,
smart quotes, typography, and mathematical symbols.
"""

import unicodedata
from typing import Optional


def normalize_unicode(text: Optional[str]) -> str:
    """Normalizes string to Unicode NFC format."""
    if text is None:
        return ""
    return unicodedata.normalize("NFC", text)


def normalize_comparison(text: str) -> str:
    """Normalizes string for fuzzy/case-insensitive comparison."""
    return normalize_unicode(text).strip().casefold()


# Vietnamese alphabet with accents for validation
VIETNAMESE_ACCENTED_CHARS = (
    "àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵ"
    "ÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴđĐ"
)


def contains_vietnamese(text: str) -> bool:
    """Returns True if string contains Vietnamese diacritics."""
    return any(c in VIETNAMESE_ACCENTED_CHARS for c in text)
