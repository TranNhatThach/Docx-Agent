import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

from docx_agent.agent import DocumentAgent
from docx_agent.media.diagrams import DiagramSynthesizer


def main():
    output_path = Path(__file__).parent / "output_03_mcp.docx"
    print("[*] Initializing Agent with new document...")
    agent = DocumentAgent()

    print("[*] Simulating AI Agent tool call: 'append_heading'...")
    agent.append("Kiến Trúc Tích Hợp MCP & AI Agent", heading_level=1)

    print("[*] Simulating AI Agent tool call: 'append_paragraph'...")
    agent.append(
        "Docx-Agent cung cấp giao thức Model Context Protocol (MCP) chuẩn hóa giúp các AI Coding Agent (Cursor, Antigravity, Claude Code) thao tác tài liệu an toàn."
    )

    print("[*] Simulating AI Agent tool call: 'generate_diagram'...")
    diag_items = [
        "LLM Agent (Antigravity / Cursor)",
        "MCP Stdio Protocol",
        "Docx-Agent Canonical Engine",
        "OpenXML Storage (.docx)",
    ]
    diag_block = DiagramSynthesizer.generate_architecture_diagram(diag_items, title="Pipeline Tích Hợp")
    print(f"[+] Generated Mermaid diagram payload: {diag_block.caption}")
    agent.append(f"Sơ đồ: {diag_block.caption}\n{diag_block.source_code}")

    print(f"[*] Saving verified document to: {output_path}")
    saved = agent.save(str(output_path), verify=True)
    print(f"[+] Saved successfully to: {saved}")


if __name__ == "__main__":
    main()
