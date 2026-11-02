#!/usr/bin/env python3
"""Direct-download page for Arena live preview deliverables."""

from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parents[1] / "deliverables"
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
FILES = {
    # Current October briefing; short aliases point to the latest deck.
    "/download/pptx": (ROOT / "October_2026_Semiconductor_News_CEO_Briefing.pptx", PPTX_MIME),
    "/download/zip": (ROOT / "October_2026_Semiconductor_News_CEO_Briefing.zip", "application/zip"),
    "/download/october/pptx": (ROOT / "October_2026_Semiconductor_News_CEO_Briefing.pptx", PPTX_MIME),
    "/download/october/zip": (ROOT / "October_2026_Semiconductor_News_CEO_Briefing.zip", "application/zip"),
    # Previous briefings remain available for reference.
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
<span class="tag">CEO BRIEFING · OCTOBER 2026</span><h1>2026년 10월 세미뉴스 종합 보고</h1>
<p>16개 이슈·258개 헤드라인을 분석해 AI 수요, 메모리 공급권, 첨단패키징, 전력 병목, 중국 추격과 CEO 90일 실행계획을 19장으로 정리했습니다.</p>
<div class="buttons"><a class="primary" href="/download/october/pptx">PPTX 직접 다운로드</a><a class="secondary" href="/download/october/zip">ZIP 파일로 다운로드</a></div>
<div class="meta"><span>16:9</span><span>19 slides</span><span>편집 가능한 도형 중심</span><span>약 96 KB</span></div>
<div class="archive"><b>이전 보고서</b><div><a class="secondary" href="/download/september/pptx">9월 PPTX</a><a class="secondary" href="/download/april/pptx">4월 PPTX</a></div></div>
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
