"""
Example 04: Academic Report Preset Application (TCVN / UTC Standards).
Demonstrates applying standardized margin presets, line spacing 1.4,
Times New Roman typography, and automated header/footer pagination.
"""

import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

from docx_agent.agent import DocumentAgent


def main():
    output_path = Path(__file__).parent / "output_04_academic_preset.docx"
    print("[*] Creating new document with Academic VN Preset...")
    agent = DocumentAgent()

    print("[*] Applying 'academic_vn' preset...")
    agent.apply_preset("academic_vn")

    print("[*] Adding cover page & thesis contents...")
    agent.append("TRƯỜNG ĐẠI HỌC GIAO THÔNG VẬN TẢI\nKHOA CÔNG NGHỆ THÔNG TIN", heading_level=1)
    agent.append("BÁO CÁO BÀI TẬP LỚN CƠ SỞ DỮ LIỆU ORACLE")
    agent.append("Học phần: Hệ quản trị cơ sở dữ liệu Oracle\nSinh viên thực hiện: Trần Nhật Thạch")

    print("[*] Configuring isolated headers, footers & page numbers...")
    agent.set_page_numbers(format_str="Trang {PAGE} / {NUMPAGES}", alignment="right")

    print(f"[*] Saving and verifying academic report to: {output_path}")
    saved = agent.save(str(output_path), verify=True)
    print(f"[+] Academic document verified and saved to: {saved}")


if __name__ == "__main__":
    main()
