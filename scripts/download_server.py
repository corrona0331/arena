#!/usr/bin/env python3
"""Direct-download page for Arena live preview deliverables."""

from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parents[1] / "deliverables"
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
FILES = {
    # Current English animated edition; short aliases point to this deck.
    "/download/pptx": (ROOT / "October_2026_Semiconductor_News_CEO_Briefing_EN_Animated.pptx", PPTX_MIME),
    "/download/zip": (ROOT / "October_2026_Semiconductor_News_CEO_Briefing_EN_Animated.zip", "application/zip"),
    "/download/october/en/pptx": (ROOT / "October_2026_Semiconductor_News_CEO_Briefing_EN_Animated.pptx", PPTX_MIME),
    "/download/october/en/zip": (ROOT / "October_2026_Semiconductor_News_CEO_Briefing_EN_Animated.zip", "application/zip"),
    # Korean October edition and previous briefings remain available.
    "/download/october/pptx": (ROOT / "October_2026_Semiconductor_News_CEO_Briefing.pptx", PPTX_MIME),
    "/download/october/zip": (ROOT / "October_2026_Semiconductor_News_CEO_Briefing.zip", "application/zip"),
    "/download/october/ko/pptx": (ROOT / "October_2026_Semiconductor_News_CEO_Briefing.pptx", PPTX_MIME),
    "/download/october/ko/zip": (ROOT / "October_2026_Semiconductor_News_CEO_Briefing.zip", "application/zip"),
    "/download/september/pptx": (ROOT / "September_2026_Semiconductor_News_CEO_Briefing.pptx", PPTX_MIME),
    "/download/september/zip": (ROOT / "September_2026_Semiconductor_News_CEO_Briefing.zip", "application/zip"),
    "/download/april/pptx": (ROOT / "April_2026_Semiconductor_News_CEO_Briefing.pptx", PPTX_MIME),
    "/download/april/zip": (ROOT / "April_2026_Semiconductor_News_CEO_Briefing.zip", "application/zip"),
}

HTML = """<!doctype html>
<html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>CEO 보고 PPT 다운로드</title>
<style>
:root{color-scheme:dark}*{box-sizing:border-box}body{margin:0;background:#07111f;color:#f7fafc;font-family:Arial,'Malgun Gothic',sans-serif;display:grid;place-items:center;min-height:100vh;padding:28px}
main{width:min(780px,94vw);background:#0d1b2a;border:1px solid #1b344b;border-radius:24px;padding:42px;box-shadow:0 20px 70px #0008}
.tag{display:inline-block;color:#2dd4bf;background:#102437;border-radius:999px;padding:8px 14px;font-size:12px;font-weight:700;letter-spacing:.06em}
h1{font-size:30px;margin:22px 0 10px}p{color:#91a4b7;line-height:1.65;margin:0 0 25px}.buttons{display:grid;grid-template-columns:1fr 1fr;gap:12px}
a{display:block;text-decoration:none;text-align:center;font-weight:800;border-radius:12px;padding:17px}
.primary{background:#2dd4bf;color:#07111f}.secondary{background:#142b40;color:#f7fafc;border:1px solid #2d4b65}
.meta{display:flex;gap:10px;flex-wrap:wrap;margin-top:20px}.meta span{color:#91a4b7;background:#102437;padding:7px 10px;border-radius:8px;font-size:12px}
.archive{border-top:1px solid #1b344b;margin-top:30px;padding-top:22px;display:flex;align-items:center;justify-content:space-between;gap:16px}.archive b{font-size:14px}.archive a{display:inline-block;padding:9px 13px;font-size:12px;margin-left:6px}
small{display:block;color:#60758a;margin-top:24px;text-align:center}@media(max-width:620px){main{padding:28px}.buttons{grid-template-columns:1fr}.archive{display:block}.archive div{margin-top:12px}}
</style></head><body><main>
<span class="tag">CEO BRIEFING · ENGLISH ANIMATED EDITION</span><h1>October 2026 Semiconductor News</h1>
<p>An English CEO briefing built from 16 issues and 258 headlines, with restrained native PowerPoint fade and wipe reveals across 19 slides.</p>
<div class="buttons"><a class="primary" href="/download/october/en/pptx">Download animated PPTX</a><a class="secondary" href="/download/october/en/zip">Download ZIP</a></div>
<div class="meta"><span>16:9</span><span>19 slides</span><span>Editable native shapes</span><span>Fade + wipe animations</span><span>117 KB</span></div>
<div class="archive"><b>Other editions</b><div><a class="secondary" href="/download/october/ko/pptx">Korean PPTX</a><a class="secondary" href="/download/september/pptx">September PPTX</a></div></div>
<small>브라우저에서 PPTX가 차단되면 ZIP을 내려받은 뒤 압축을 해제하세요.</small>
</main></body></html>""".encode("utf-8")


class Handler(BaseHTTPRequestHandler):
    def _serve(self, head_only: bool = False):
        path = urlparse(self.path).path
        if path in ("/", "/index.html"):
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(HTML)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.end_headers()
            if not head_only:
                self.wfile.write(HTML)
            return

        item = FILES.get(path)
        if item:
            file_path, mime = item
            if not file_path.is_file():
                self.send_error(404, "Deliverable not found")
                return
            size = file_path.stat().st_size
            self.send_response(200)
            self.send_header("Content-Type", mime)
            self.send_header("Content-Length", str(size))
            self.send_header("Content-Disposition", f'attachment; filename="{file_path.name}"')
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.end_headers()
            if not head_only:
                self.wfile.write(file_path.read_bytes())
            return

        self.send_error(404)

    def do_GET(self):
        self._serve(False)

    def do_HEAD(self):
        self._serve(True)

    def log_message(self, fmt, *args):
        print(f"{self.client_address[0]} - {fmt % args}", flush=True)


if __name__ == "__main__":
    print("PPT download page listening on 0.0.0.0:8000", flush=True)
    ThreadingHTTPServer(("0.0.0.0", 8000), Handler).serve_forever()
