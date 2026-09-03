import http.server
import socketserver
import os
import subprocess

PORT = 8000
PPTX_FILE = "Test.pptx"  # Your PowerPoint file name

class LaunchHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path == "/start-presentation":
            # Launch PowerPoint directly in Full-Screen Slide Show mode (/s flag)
            abs_path = os.path.abspath(PPTX_FILE)
            subprocess.Popen(["powerpnt.exe", "/s", abs_path])
            
            self.send_response(200)
            self.send_header("Content-Type", "text/plain")
            self.end_headers()
            self.wfile.write(b"OK")
            return
            
        super().do_GET()

print(f"Server running on http://localhost:{PORT}")
with socketserver.TCPServer(("", PORT), LaunchHandler) as httpd:
    httpd.serve_forever()