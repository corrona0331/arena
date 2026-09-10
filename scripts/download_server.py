#!/usr/bin/env python3
"""Temporary file-download page for Arena live preview."""

from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parents[1] / "deliverables"
FILES = {
    "/download/pptx": (ROOT / "April_2026_Semiconductor_News_CEO_Briefing.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
    "/download/zip": (ROOT / "April_2026_Semiconductor_News_CEO_Briefing.zip", "application/zip"),
}

HTML = """<!doctype html>
<html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>CEO 보고 PPT 다운로드</title>
<style>
:root{color-scheme:dark} body{margin:0;background:#07111f;color:#f7fafc;font-family:Arial,'Malgun Gothic',sans-serif;display:grid;place-items:center;min-height:100vh}
main{width:min(720px,88vw);background:#0d1b2a;border:1px solid #1b344b;border-radius:24px;padding:42px;box-shadow:0 20px 70px #0008}
.tag{display:inline-block;color:#2dd4bf;background:#102437;border-radius:999px;padding:8px 14px;font-size:12px;font-weight:700;letter-spacing:.06em}
h1{font-size:30px;margin:22px 0 10px} p{color:#91a4b7;line-height:1.65;margin:0 0 28px}
a{display:block;text-decoration:none;text-align:center;font-weight:800;border-radius:12px;padding:17px;margin:12px 0}
.primary{background:#2dd4bf;color:#07111f}.secondary{background:#142b40;color:#f7fafc;border:1px solid #2d4b65}
small{display:block;color:#60758a;margin-top:24px;text-align:center}
</style></head><body><main>
<span class="tag">CEO BRIEFING</span><h1>2026년 4월 세미뉴스 종합 보고</h1>
<p>아래 첫 번째 버튼으로 PPTX를 바로 내려받으세요. 브라우저에서 차단될 경우 ZIP 버튼을 이용한 뒤 압축을 해제하면 됩니다.</p>
<a class="primary" href="/download/pptx">PPTX 직접 다운로드</a>
<a class="secondary" href="/download/zip">ZIP 파일로 다운로드</a>
<small>PPTX 78 KB · ZIP 68 KB · 14 slides</small>
</main></body></html>""".encode("utf-8")


class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        path = urlparse(self.path).path
        if path in ("/", "/index.html"):
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(HTML)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(HTML)
            return
        item = FILES.get(path)
        if item:
            file_path, mime = item
            data = file_path.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", mime)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Content-Disposition", f'attachment; filename="{file_path.name}"')
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(data)
            return
        self.send_error(404)

    def log_message(self, fmt, *args):
        print(f"{self.client_address[0]} - {fmt % args}", flush=True)


if __name__ == "__main__":
    print("PPT download page listening on 0.0.0.0:8000", flush=True)
    ThreadingHTTPServer(("0.0.0.0", 8000), Handler).serve_forever()
