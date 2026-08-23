"""
Local Workspace HTTP Server for Visual Document Workspace.
"""

import http.server
import socketserver
import webbrowser
from pathlib import Path
from typing import Optional

WORKSPACE_DIR = Path(__file__).parent


class WorkspaceServer:
    """
    Hosts the visual document workspace locally on demand.
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
                else:
                    super().do_GET()

        with socketserver.TCPServer(("", port), Handler) as httpd:
            url = f"http://localhost:{port}"
            print(f"Docx-Agent V2 Workspace running at: {url}")
            if open_browser:
                webbrowser.open(url)
            try:
                httpd.serve_forever()
            except KeyboardInterrupt:
                print("\nWorkspace server stopped.")
