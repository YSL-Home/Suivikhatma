#!/usr/bin/env python3
# Serveur web minimaliste pour suivikhatma
import http.server, os, socketserver

PORT = int(os.environ.get("PORT", 8080))
os.chdir(os.path.dirname(os.path.abspath(__file__)))

class Handler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path == "/" or self.path == "":
            self.path = "/index.html"
        return super().do_GET()
    def log_message(self, format, *args):
        pass  # silence logs

print(f"✅ suivikhatma en ligne sur le port {PORT}")
with socketserver.TCPServer(("", PORT), Handler) as httpd:
    httpd.serve_forever()
