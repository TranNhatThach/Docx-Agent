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
                    with open(html_file, "r", encoding="utf-8") as f:
                        html_text = f.read()
                    jszip_f = WORKSPACE_DIR / "jszip.min.js"
                    docx_f = WORKSPACE_DIR / "docx-preview.min.js"
                    if jszip_f.exists():
                        html_text = html_text.replace('<script src="https://unpkg.com/jszip/dist/jszip.min.js"></script>', f'<script>{jszip_f.read_text(encoding="utf-8")}</script>')
                    if docx_f.exists():
                        html_text = html_text.replace('<script src="https://cdn.jsdelivr.net/npm/docx-preview@0.4.0/dist/docx-preview.min.js"></script>', f'<script>{docx_f.read_text(encoding="utf-8")}</script>')
                    self.wfile.write(html_text.encode("utf-8"))
                elif self.path == "/api/document":
                    self.send_response(200)
                    self.send_header("Content-type", "application/json; charset=utf-8")
                    self.end_headers()
                    try:
                        if file_path and Path(file_path).exists():
                            import base64
                            doc_data = WorkspaceBridge.load_document_payload(file_path)
                            with open(file_path, "rb") as bf:
                                docx_b64 = base64.b64encode(bf.read()).decode("ascii")
                            response_data = {
                                "payload": doc_data,
                                "docxBase64": docx_b64
                            }
                        else:
                            response_data = {"error": "Không có tệp tài liệu được chỉ định."}
                    except Exception as e:
                        response_data = {"error": str(e)}
                    self.wfile.write(json.dumps(response_data, ensure_ascii=False).encode("utf-8"))
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

        socketserver.ThreadingTCPServer.allow_reuse_address = True
        with socketserver.ThreadingTCPServer(("", port), Handler) as httpd:
            url = f"http://localhost:{port}"
            print(f"Docx-Agent V2.1 Workspace running at: {url}")
            if open_browser:
                webbrowser.open(url)
            try:
                httpd.serve_forever()
            except KeyboardInterrupt:
                print("\nWorkspace server stopped.")
