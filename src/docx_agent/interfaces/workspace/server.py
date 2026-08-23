"""
Local Workspace HTTP Server for Visual Document Workspace with Live Document Ingestion.
"""

import http.server
import socketserver
import webbrowser
import json
from pathlib import Path
from typing import Optional
from docx_agent.adapters.docx import DocxImporter
from docx_agent.canonical.model import HeadingBlock, ParagraphBlock, TableBlock

WORKSPACE_DIR = Path(__file__).parent


class WorkspaceServer:
    """
    Hosts the visual document workspace locally on demand with live document API.
    """

    @staticmethod
    def launch(file_path: Optional[str] = None, port: int = 8765, open_browser: bool = True) -> None:
        html_file = WORKSPACE_DIR / "app.html"
        
        doc_data = None
        if file_path and Path(file_path).exists():
            try:
                doc_node = DocxImporter.import_docx(file_path)
                headings = []
                body_html_parts = []
                for sec in doc_node.sections:
                    for blk in sec.blocks:
                        if isinstance(blk, HeadingBlock):
                            headings.append({"level": blk.level, "text": blk.full_text, "id": blk.id})
                            body_html_parts.append(f"<h{blk.level} id='{blk.id}'>{blk.full_text}</h{blk.level}>")
                        elif isinstance(blk, ParagraphBlock):
                            if blk.full_text.strip():
                                body_html_parts.append(f"<p id='{blk.id}'>{blk.full_text}</p>")
                        elif isinstance(blk, TableBlock):
                            tbl_rows = "".join(f"<tr>{''.join(f'<td>{c.text}</td>' for c in row)}</tr>" for row in blk.cells)
                            body_html_parts.append(f"<table id='{blk.id}'>{tbl_rows}</table>")
                doc_data = {
                    "title": doc_node.title or Path(file_path).stem,
                    "headings": headings,
                    "body_html": "".join(body_html_parts),
                }
            except Exception as e:
                doc_data = {"error": str(e)}

        class Handler(http.server.SimpleHTTPRequestHandler):
            def do_GET(self):
                if self.path in ("/", "/index.html"):
                    self.send_response(200)
                    self.send_header("Content-type", "text/html; charset=utf-8")
                    self.end_headers()
                    with open(html_file, "rb") as f:
                        self.wfile.write(f.read())
                elif self.path == "/api/document":
                    self.send_response(200)
                    self.send_header("Content-type", "application/json; charset=utf-8")
                    self.end_headers()
                    self.wfile.write(json.dumps(doc_data or {}, ensure_ascii=False).encode("utf-8"))
                else:
                    super().do_GET()

        socketserver.TCPServer.allow_reuse_address = True
        with socketserver.TCPServer(("", port), Handler) as httpd:
            url = f"http://localhost:{port}"
            print(f"Docx-Agent V2 Workspace running at: {url}")
            if open_browser:
                webbrowser.open(url)
            try:
                httpd.serve_forever()
            except KeyboardInterrupt:
                print("\nWorkspace server stopped.")
