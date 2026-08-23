"""
Citation Formatter: Generates standard in-text citations and bibliographies (APA, IEEE, Academic-VN).
"""

from typing import List
from docx_agent.canonical.model import SourceMetadata


def _extract_last_name(author_str: str) -> str:
    """Extracts author last name regardless of 'Lastname, First' or 'First Lastname' format."""
    clean = author_str.strip()
    if "," in clean:
        return clean.split(",")[0].strip()
    parts = clean.split()
    return parts[-1].strip() if parts else "Anon"


class CitationFormatter:
    """
    Renders academic citations and reference list entries adhering strictly to style manuals.
    """

    @staticmethod
    def format_intext(source: SourceMetadata, style: str = "apa", number: int = 1) -> str:
        """Formats in-text citation marker."""
        style_clean = style.lower().replace("-", "_")

        if style_clean == "ieee":
            return f"[{number}]"

        # Author-date styles (APA, Academic-VN)
        authors = source.authors
        author_str = "Anon"
        if len(authors) == 1:
            author_str = _extract_last_name(authors[0])
        elif len(authors) == 2:
            author_str = f"{_extract_last_name(authors[0])} & {_extract_last_name(authors[1])}"
        elif len(authors) > 2:
            author_str = f"{_extract_last_name(authors[0])} et al."

        year_str = str(source.year) if source.year else "n.d."
        return f"({author_str}, {year_str})"

    @staticmethod
    def format_bibliography_entry(source: SourceMetadata, style: str = "apa", number: int = 1) -> str:
        """Formats full reference list entry."""
        style_clean = style.lower().replace("-", "_")
        authors_str = ", ".join(source.authors) if source.authors else "Anonymous"
        year_str = f"({source.year})" if source.year else "(n.d.)"
        title_str = source.title
        pub_str = source.publication or ""
        doi_str = f"https://doi.org/{source.doi}" if source.doi else (source.url or "")

        if style_clean == "ieee":
            return f"[{number}] {authors_str}, \"{title_str},\" {pub_str}, {source.year or ''}. {doi_str}".strip()
        elif style_clean == "academic_vn":
            # Vietnamese academic reference format
            return f"{authors_str} {year_str}, \"{title_str}\", {pub_str}, {doi_str}".strip()
        else:
            # Default: APA 7th
            return f"{authors_str} {year_str}. {title_str}. {pub_str}. {doi_str}".strip()
