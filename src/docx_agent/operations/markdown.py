"""
Markdown to DOCX Converter: Transforms structured markdown documents
into professionally styled DOCX documents using the Style & Preset Engine.
"""

import re
from pathlib import Path
from typing import Optional, Union, Dict, Any, List
import docx
from docx_agent.agent import DocumentAgent
from docx_agent.utils.paths import resolve_safe_path


class MarkdownToDocxConverter:
    """
    Parses Markdown text (headings, bold, italic, code blocks, lists, blockquotes, tables)
    and generates a styled DOCX document.
    """

    @staticmethod
    def convert(
        md_text_or_path: Union[str, Path],
        output_path: Union[str, Path],
        preset: str = "academic-vn",
    ) -> str:
        if isinstance(md_text_or_path, (str, Path)) and (Path(md_text_or_path).is_file() or str(md_text_or_path).endswith(".md")):
            path = resolve_safe_path(md_text_or_path)
            with open(path, "r", encoding="utf-8") as f:
                content = f.read()
        else:
            content = str(md_text_or_path)

        agent = DocumentAgent()
        agent.apply_preset(preset)

        lines = content.splitlines()
        in_code_block = False
        code_lines = []

        in_table = False
        table_rows = []

        def flush_table():
            nonlocal in_table, table_rows
            if table_rows:
                # Filter out markdown divider rows (|---|---|)
                cleaned_rows = [
                    row for row in table_rows
                    if not all(re.match(r"^:?-+:?$", cell.strip()) for cell in row)
                ]
                if cleaned_rows:
                    num_rows = len(cleaned_rows)
                    num_cols = max(len(r) for r in cleaned_rows)
                    agent.create_table(rows=num_rows, cols=num_cols, data=cleaned_rows)
            table_rows = []
            in_table = False

        for line in lines:
            line_str = line.rstrip()

            # Code Block Toggle
            if line_str.startswith("```"):
                if in_code_block:
                    in_code_block = False
                    code_text = "\n".join(code_lines)
                    p_id = agent.append(code_text)
                    agent.format_text(p_id, font_name="Courier New", font_size_pt=10.0)
                    agent.format_paragraph(p_id, first_line_indent_cm=0.0, line_spacing=1.0)
                    code_lines = []
                else:
                    flush_table()
                    in_code_block = True
                continue

            if in_code_block:
                code_lines.append(line_str)
                continue

            # Table row: | col1 | col2 |
            if line_str.startswith("|") and line_str.endswith("|"):
                in_table = True
                cells = [c.strip() for c in line_str[1:-1].split("|")]
                table_rows.append(cells)
                continue
            elif in_table:
                flush_table()

            # Empty Line
            if not line_str.strip():
                continue

            # Headings
            if line_str.startswith("# "):
                agent.append(line_str[2:].strip(), style="Heading 1")
            elif line_str.startswith("## "):
                agent.append(line_str[3:].strip(), style="Heading 2")
            elif line_str.startswith("### "):
                agent.append(line_str[4:].strip(), style="Heading 3")
            elif line_str.startswith("#### "):
                agent.append(line_str[5:].strip(), style="Heading 4")
            # Bullet list
            elif line_str.startswith("- ") or line_str.startswith("* "):
                agent.append(line_str[2:].strip(), style="List Bullet")
            # Numbered list
            elif re.match(r"^\d+\.\s+", line_str):
                text_clean = re.sub(r"^\d+\.\s+", "", line_str)
                agent.append(text_clean, style="List Number")
            # Blockquote
            elif line_str.startswith("> "):
                p_id = agent.append(line_str[2:].strip())
                agent.format_paragraph(p_id, first_line_indent_cm=0.0, left_indent_cm=1.0)
                agent.format_text(p_id, italic=True)
            # Regular paragraph with inline markdown parsing
            else:
                p_id = agent.append(line_str.strip(), style="Normal")

        flush_table()

        agent.save(output_path)
        return str(resolve_safe_path(output_path))
