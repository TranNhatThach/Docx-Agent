"""
Local Workspace HTTP Server for Visual Document Workspace with Live Document Ingestion & Save API.
"""

import http.server
import socketserver
import webbrowser
import json
from pathlib import Path
from typing import Optional

from docx_agent.interfaces.workspace.bridge import WorkspaceBridge

WORKSPACE_DIR = Path(__file__).parent


class WorkspaceServer:
    """
    Hosts the visual document workspace locally on demand with live document & save API.
    """

    @staticmethod
    def launch(file_path: Optional[str] = None, port: int = 8765, open_browser: bool = True) -> None:
        html_file = WORKSPACE_DIR / "app.html"

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
                    try:
                        if file_path and Path(file_path).exists():
                            doc_data = WorkspaceBridge.load_document_payload(file_path)
                        else:
                            doc_data = {"error": "Không có tệp tài liệu được chỉ định."}
                    except Exception as e:
                        doc_data = {"error": str(e)}
                    self.wfile.write(json.dumps(doc_data, ensure_ascii=False).encode("utf-8"))
                else:
                    super().do_GET()

            def do_POST(self):
                if self.path == "/api/save":
                    content_length = int(self.headers.get("Content-Length", 0))
                    body = self.rfile.read(content_length).decode("utf-8")
                    try:
                        save_payload = json.loads(body)
                        res = WorkspaceBridge.save_document_payload(
                            file_path=file_path or "output.docx",
                            document_data=save_payload,
                        )
                        self.send_response(200)
                        self.send_header("Content-type", "application/json; charset=utf-8")
                        self.end_headers()
                        self.wfile.write(json.dumps(res, ensure_ascii=False).encode("utf-8"))
                    except Exception as e:
                        self.send_response(500)
                        self.send_header("Content-type", "application/json; charset=utf-8")
                        self.end_headers()
                        self.wfile.write(json.dumps({"success": False, "error": str(e)}).encode("utf-8"))
                else:
                    self.send_response(404)
                    self.end_headers()

        socketserver.TCPServer.allow_reuse_address = True
        with socketserver.TCPServer(("", port), Handler) as httpd:
            url = f"http://localhost:{port}"
            print(f"Docx-Agent V2.1 Workspace running at: {url}")
            if open_browser:
                webbrowser.open(url)
            try:
                httpd.serve_forever()
            except KeyboardInterrupt:
                print("\nWorkspace server stopped.")
