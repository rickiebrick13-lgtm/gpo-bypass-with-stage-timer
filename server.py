import http.server
import socketserver
import os
import subprocess

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PPTX_FILENAME = "Test.pptx"
PPTX_PATH = os.path.join(BASE_DIR, PPTX_FILENAME)
PORT = 8000

class LaunchHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path == "/start-presentation":
            if os.path.exists(PPTX_PATH):
                print(f"\n[SUCCESS] Launching Fullscreen PowerPoint: {PPTX_PATH}")
                
                # 'start /MAX' forces the window to open maximized over the browser
                cmd_command = f'cmd /c start "" /MAX powerpnt /s "{PPTX_PATH}"'
                subprocess.Popen(cmd_command, shell=True)
                
                self.send_response(200)
                self.send_header("Content-Type", "text/plain")
                self.end_headers()
                self.wfile.write(b"OK")
            else:
                print(f"\n[ERROR] File missing! Looked for: {PPTX_PATH}")
                self.send_response(404)
                self.end_headers()
                self.wfile.write(b"File Not Found")
            return
            
        super().do_GET()

socketserver.TCPServer.allow_reuse_address = True

print("==================================================")
print(f"Stage Clock Server Running on: http://localhost:{PORT}")
print(f"Directory: {BASE_DIR}")
print(f"Target PPTX: {PPTX_PATH}")
print(f"File Exists? -> {os.path.exists(PPTX_PATH)}")
print("==================================================")

try:
    with socketserver.TCPServer(("", PORT), LaunchHandler) as httpd:
        httpd.serve_forever()
except KeyboardInterrupt:
    print("\n[INFO] Server shut down successfully.")