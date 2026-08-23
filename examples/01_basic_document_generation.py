"""
Example 01: Basic Document Generation with Docx-Agent.
Demonstrates programmatic document creation, typography formatting, and verified export.
"""

from pathlib import Path
from docx_agent.agent import DocumentAgent


def main():
    output_path = Path(__file__).parent / "output_01_basic.docx"
    print("[*] Initializing DocumentAgent...")
    agent = DocumentAgent()

    print("[*] Appending headings and formatted paragraphs...")
    agent.append("1. Tổng Quan Về Kiến Trúc Oracle 19c", heading_level=1)
    agent.append(
        "Hệ quản trị cơ sở dữ liệu Oracle là một trong những hệ thống RDBMS hàng đầu thế giới với kiến trúc Multitenant mạnh mẽ."
    )
    agent.append("1.1. Cấu Trúc Bộ Nhớ SGA và PGA", heading_level=2)
    agent.append(
        "System Global Area (SGA) là vùng nhớ chia sẻ chung chứa Database Buffer Cache, Shared Pool và Redo Log Buffer."
    )

    print("[*] Creating a formatted table with structured data...")
    table_data = [
        ["Thành phần bộ nhớ", "Mục đích sử dụng"],
        ["Database Buffer Cache", "Lưu trữ bản sao các khối dữ liệu đọc từ đĩa"],
        ["Shared Pool", "Lưu trữ mã SQL đã parse và Execution Plans"],
    ]
    tbl_id = agent.create_table(
        rows=3,
        cols=2,
        data=table_data,
        repeat_header=True,
        col_widths_cm=[5.0, 11.0],
    )
    print(f"[+] Created table with ID: {tbl_id}")

    print(f"[*] Saving and verifying document to: {output_path}")
    saved_file = agent.save(str(output_path), verify=True)
    print(f"[+] Saved successfully to: {saved_file}")


if __name__ == "__main__":
    main()
