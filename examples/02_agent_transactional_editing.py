"""
Example 02: Transactional Document Editing with Rollback.
Demonstrates safe run surgery replacements, multi-step transaction staging, and undo/rollback.
"""

import importlib.util
from pathlib import Path
from docx_agent.agent import DocumentAgent
from docx_agent.transactions.transaction import TransactionContext


def main():
    sample_file = Path(__file__).parent / "output_01_basic.docx"
    if not sample_file.exists():
        print("[!] Generating base document first...")
        spec = importlib.util.spec_from_file_location("basic_gen", Path(__file__).parent / "01_basic_document_generation.py")
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        mod.main()

    print(f"[*] Opening document in TransactionContext: {sample_file}")
    with TransactionContext(sample_file, auto_backup=True) as tx:
        agent = DocumentAgent(tx.file_path)

        print("[*] Inspecting document structure...")
        summary = agent.inspect()
        print(f"[+] Total paragraphs: {summary.paragraphs_count}, Total tables: {summary.tables_count}")

        print("[*] Performing safe Run Surgery text replacement...")
        replaced = agent.replace(
            target="Hệ thống RDBMS hàng đầu",
            replacement="Hệ quản trị cơ sở dữ liệu quan hệ doanh nghiệp hàng đầu",
        )
        print(f"[+] Replaced occurrences: {replaced}")

        print("[*] Formatting paragraph properties...")
        paras = agent.find(text="System Global Area")
        if paras:
            target_id = paras[0]["id"]
            agent.format_paragraph(target_id, alignment="justify", line_spacing=1.4)
            print(f"[+] Formatted paragraph {target_id} with line_spacing 1.4")

        print("[*] Saving and verifying transaction atomically...")
        agent.save(tx.file_path, verify=True)
        print("[+] Transaction committed successfully with 100% integrity verified!")


if __name__ == "__main__":
    main()
