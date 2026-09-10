#!/usr/bin/env python3
"""Build a CEO briefing deck from the five April 2026 semiconductor-news issues.

The deck intentionally uses native PowerPoint shapes and text so that every
visual remains editable by the recipient.
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable, Sequence

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


# ---------------------------------------------------------------------------
# Theme
# ---------------------------------------------------------------------------

W = 13.333
H = 7.5

BG = "07111F"
BG_ALT = "091726"
PANEL = "0D1B2A"
PANEL_2 = "102437"
PANEL_3 = "142B40"
GRID = "1B344B"
WHITE = "F7FAFC"
INK = "DDE7F0"
MUTED = "91A4B7"
DIM = "60758A"
CYAN = "2DD4BF"
BLUE = "38BDF8"
LIME = "A3E635"
ORANGE = "FB923C"
YELLOW = "FBBF24"
RED = "F87171"
PURPLE = "A78BFA"

FONT_KR = "맑은 고딕"
FONT_EN = "Aptos"

OUT = Path("deliverables/2026_04_세미뉴스_CEO_보고.pptx")


def rgb(hex_color: str) -> RGBColor:
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def set_fill(shape, color: str | None) -> None:
    if color is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(color)


def set_line(shape, color: str | None, width: float = 1.0) -> None:
    if color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = rgb(color)
        shape.line.width = Pt(width)


def add_box(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    fill: str | None = PANEL,
    line: str | None = None,
    radius: bool = True,
    width: float = 1.0,
):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    set_fill(shape, fill)
    set_line(shape, line, width)
    return shape


def add_circle(
    slide,
    x: float,
    y: float,
    d: float,
    *,
    fill: str | None = PANEL,
    line: str | None = None,
    width: float = 1.0,
):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    set_fill(shape, fill)
    set_line(shape, line, width)
    return shape


def add_line(slide, x1: float, y1: float, x2: float, y2: float, color: str = GRID, width: float = 1.0):
    line = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(x1),
        Inches(y1),
        Inches(x2),
        Inches(y2),
    )
    set_line(line, color, width)
    return line


def add_text(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    text: str,
    *,
    size: float = 18,
    color: str = INK,
    bold: bool = False,
    font: str = FONT_KR,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    valign: MSO_ANCHOR = MSO_ANCHOR.TOP,
    margin: float = 0.0,
    line_spacing: float | None = None,
    break_line: bool = True,
):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin)
    tf.margin_bottom = Inches(margin)
    p = tf.paragraphs[0]
    p.alignment = align
    if line_spacing is not None:
        p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    return box


def add_rich_text(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    runs: Sequence[dict],
    *,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    valign: MSO_ANCHOR = MSO_ANCHOR.TOP,
    margin: float = 0.0,
    line_spacing: float | None = None,
):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin)
    tf.margin_bottom = Inches(margin)
    p = tf.paragraphs[0]
    p.alignment = align
    if line_spacing is not None:
        p.line_spacing = line_spacing
    for spec in runs:
        r = p.add_run()
        r.text = spec.get("text", "")
        r.font.name = spec.get("font", FONT_KR)
        r.font.size = Pt(spec.get("size", 18))
        r.font.bold = spec.get("bold", False)
        r.font.color.rgb = rgb(spec.get("color", INK))
    return box


def add_multiline(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    lines: Sequence[dict],
    *,
    margin: float = 0.0,
    valign: MSO_ANCHOR = MSO_ANCHOR.TOP,
):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin)
    tf.margin_bottom = Inches(margin)
    for idx, spec in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = spec.get("align", PP_ALIGN.LEFT)
        p.level = spec.get("level", 0)
        p.space_before = Pt(spec.get("before", 0))
        p.space_after = Pt(spec.get("after", 0))
        p.line_spacing = spec.get("spacing", 1.05)
        r = p.add_run()
        r.text = spec.get("text", "")
        r.font.name = spec.get("font", FONT_KR)
        r.font.size = Pt(spec.get("size", 16))
        r.font.bold = spec.get("bold", False)
        r.font.color.rgb = rgb(spec.get("color", INK))
    return box


def add_pill(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    text: str,
    *,
    fill: str = PANEL_3,
    color: str = CYAN,
    size: float = 9.5,
    line: str | None = None,
):
    add_box(slide, x, y, w, h, fill=fill, line=line)
    add_text(
        slide,
        x,
        y + 0.005,
        w,
        h - 0.01,
        text,
        size=size,
        color=color,
        bold=True,
        align=PP_ALIGN.CENTER,
        valign=MSO_ANCHOR.MIDDLE,
    )


def add_number_badge(slide, x: float, y: float, text: str, color: str = CYAN, d: float = 0.38):
    add_circle(slide, x, y, d, fill=color, line=None)
    add_text(
        slide,
        x,
        y,
        d,
        d,
        text,
        size=10.5,
        color=BG,
        bold=True,
        font=FONT_EN,
        align=PP_ALIGN.CENTER,
        valign=MSO_ANCHOR.MIDDLE,
    )


def add_header(slide, number: int, eyebrow: str, title: str, subtitle: str | None = None):
    add_pill(slide, 0.65, 0.40, 1.35, 0.28, eyebrow, fill=PANEL_2, color=CYAN, size=9)
    add_text(slide, 0.65, 0.78, 11.9, 0.42, title, size=24.5, color=WHITE, bold=True)
    if subtitle:
        add_text(slide, 0.65, 1.18, 11.9, 0.28, subtitle, size=10.5, color=MUTED)
    # top-right section index
    add_text(
        slide,
        12.15,
        0.40,
        0.50,
        0.28,
        f"{number:02d}",
        size=10,
        color=DIM,
        bold=True,
        font=FONT_EN,
        align=PP_ALIGN.RIGHT,
        valign=MSO_ANCHOR.MIDDLE,
    )


def add_footer(slide, number: int, source: str = "자료: 세미뉴스 2026년 4월호 [S1–S5]"):
    add_line(slide, 0.65, 7.09, 12.68, 7.09, GRID, 0.7)
    add_text(slide, 0.65, 7.14, 7.8, 0.18, source, size=7.4, color=DIM)
    add_text(
        slide,
        9.4,
        7.14,
        3.28,
        0.18,
        f"CONFIDENTIAL · CEO BRIEFING   {number:02d}",
        size=7.4,
        color=DIM,
        font=FONT_EN,
        align=PP_ALIGN.RIGHT,
    )


def add_slide_bg(slide, alt: bool = False):
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = rgb(BG_ALT if alt else BG)


# ---------------------------------------------------------------------------
# Slides
# ---------------------------------------------------------------------------


def slide_cover(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)

    # subtle technical grid
    for x in [8.15, 8.75, 9.35, 9.95, 10.55, 11.15, 11.75, 12.35]:
        add_line(slide, x, 0.55, x, 6.95, GRID, 0.5)
    for y in [0.75, 1.35, 1.95, 2.55, 3.15, 3.75, 4.35, 4.95, 5.55, 6.15, 6.75]:
        add_line(slide, 7.75, y, 12.75, y, GRID, 0.5)

    add_pill(slide, 0.75, 0.62, 2.55, 0.32, "CEO BRIEFING · CONFIDENTIAL", fill=PANEL_2, color=CYAN, size=9.5)
    add_text(slide, 0.75, 1.42, 6.75, 0.58, "2026년 4월 세미뉴스", size=32, color=WHITE, bold=True)
    add_text(slide, 0.75, 2.02, 6.85, 1.18, "AI가 만든 ‘구조적 슈퍼사이클’\n그리고 이동하는 병목", size=29, color=WHITE, bold=True)
    add_line(slide, 0.75, 3.45, 2.25, 3.45, CYAN, 4)
    add_text(
        slide,
        0.75,
        3.72,
        6.30,
        0.94,
        "5개 이슈의 핵심 신호를 시장·기술·경쟁·공급망 관점으로 통합하고\nCEO가 즉시 결정해야 할 우선순위를 제시합니다.",
        size=14.3,
        color=MUTED,
    )
    add_text(slide, 0.75, 6.40, 5.8, 0.25, "SEMI NEWS  |  APR 07 · 08 · 14 · 22 · 28", size=10.5, color=CYAN, bold=True, font=FONT_EN)
    add_text(slide, 0.75, 6.70, 5.8, 0.20, "Prepared for CEO · 2026.04", size=9, color=DIM, font=FONT_EN)

    # Stylized editable chip / wafer illustration
    cx, cy = 10.25, 3.55
    for d, c, lw in [(3.9, GRID, 1.2), (3.2, PANEL_3, 1.0), (2.45, BLUE, 1.3)]:
        add_circle(slide, cx - d / 2, cy - d / 2, d, fill=None, line=c, width=lw)
    chip = add_box(slide, 9.15, 2.46, 2.18, 2.18, fill=PANEL_2, line=CYAN, radius=True, width=2)
    chip.rotation = 45
    core = add_box(slide, 9.72, 3.03, 1.04, 1.04, fill=CYAN, line=None, radius=False)
    core.rotation = 45
    add_text(slide, 9.46, 3.23, 1.56, 0.40, "AI", size=22, color=BG, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    nodes = [
        (8.08, 1.40, BLUE), (9.18, 0.91, CYAN), (11.60, 1.28, LIME), (12.14, 2.65, ORANGE),
        (12.05, 4.70, YELLOW), (11.05, 5.73, PURPLE), (8.84, 5.80, BLUE), (7.72, 4.44, CYAN),
    ]
    for nx, ny, c in nodes:
        add_circle(slide, nx, ny, 0.16, fill=c, line=None)
        add_line(slide, nx + 0.08, ny + 0.08, cx, cy, c, 0.9)
    add_pill(slide, 8.35, 6.30, 4.20, 0.32, "DEMAND → BOTTLENECK → VALUE MIGRATION", fill=PANEL_2, color=BLUE, size=9.2)
    return slide


def slide_executive_summary(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 2, "EXECUTIVE", "4월의 메시지: 호황이 아니라 ‘산업 운영모델’이 바뀌고 있다")

    # Main thesis
    add_box(slide, 0.65, 1.48, 7.65, 2.17, fill=PANEL, line=GRID)
    add_pill(slide, 0.92, 1.74, 1.18, 0.30, "ONE-LINE", fill=PANEL_3, color=LIME, size=9)
    add_rich_text(
        slide, 0.92, 2.16, 7.05, 0.82,
        [
            {"text": "AI 수요가 ", "size": 23, "color": WHITE, "bold": True},
            {"text": "메모리·패키징·기판·소재·전력", "size": 23, "color": CYAN, "bold": True},
            {"text": "을 동시에 병목 자산으로 바꾸고 있다.", "size": 23, "color": WHITE, "bold": True},
        ],
    )
    add_text(slide, 0.92, 3.05, 6.95, 0.34, "따라서 승부는 ‘단일 칩’이 아니라 고객 락인·생산속도·공급망 복원력의 결합에서 결정됩니다.", size=11.5, color=MUTED)

    # CEO imperative
    add_box(slide, 8.55, 1.48, 4.13, 2.17, fill=PANEL_2, line=CYAN, width=1.3)
    add_pill(slide, 8.84, 1.74, 1.53, 0.30, "CEO IMPERATIVE", fill=CYAN, color=BG, size=9)
    add_multiline(slide, 8.84, 2.17, 3.48, 1.12, [
        {"text": "성장 기회는 선점하되", "size": 18, "color": WHITE, "bold": True},
        {"text": "① 장기계약으로 수요를 고정하고", "size": 11.6, "color": INK, "before": 8},
        {"text": "② 병목 역량에 CAPEX를 집중하며", "size": 11.6, "color": INK, "before": 4},
        {"text": "③ 소재·중국 리스크를 동시에 헤지", "size": 11.6, "color": INK, "before": 4},
    ])

    # Metrics
    metrics = [
        ("5년+", "HBM 공급부족\n지속 전망", CYAN, "[S1]"),
        ("$135.1B", "글로벌 장비 매출\n전년비 +15%", BLUE, "[S2]"),
        ("+55%", "테스트 장비\n매출 증가", LIME, "[S2]"),
        ("65%", "국내 헬륨의\n카타르 의존", ORANGE, "[S1]"),
    ]
    x0 = 0.65
    card_w = 2.86
    for i, (big, label, c, src) in enumerate(metrics):
        x = x0 + i * 3.01
        add_box(slide, x, 3.94, card_w, 1.60, fill=PANEL, line=GRID)
        add_line(slide, x + 0.20, 4.18, x + 0.20, 5.27, c, 3)
        add_text(slide, x + 0.40, 4.12, 2.15, 0.44, big, size=22, color=c, bold=True, font=FONT_EN)
        add_text(slide, x + 0.40, 4.62, 2.12, 0.52, label, size=10.7, color=INK, bold=True)
        add_text(slide, x + 2.35, 5.17, 0.25, 0.14, src, size=7, color=DIM, align=PP_ALIGN.RIGHT)

    # Bottom take-away strip
    add_box(slide, 0.65, 5.82, 12.03, 0.82, fill=PANEL_3, line=None)
    add_pill(slide, 0.88, 6.06, 1.40, 0.30, "CEO TAKEAWAY", fill=CYAN, color=BG, size=9)
    add_text(slide, 2.52, 5.98, 9.78, 0.42, "‘공격적 증설’보다 중요한 것은 고객 커버리지·수율·핵심소재를 묶은 단계별 투자 규율입니다.", size=14.5, color=WHITE, bold=True, valign=MSO_ANCHOR.MIDDLE)

    add_footer(slide, 2, "자료: [S1] Apr 07, [S2] Apr 08, [S3] Apr 14, [S4] Apr 22, [S5] Apr 28")
    return slide


def slide_timeline(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 3, "SIGNAL TIMELINE", "5개 이슈를 관통한 신호: 병목의 연쇄 이동", "각 날짜의 개별 뉴스가 월말에는 하나의 구조적 흐름으로 수렴")

    # central line and directional arrow
    add_line(slide, 1.00, 3.62, 12.05, 3.62, GRID, 5)
    chevron = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(11.93), Inches(3.43), Inches(0.55), Inches(0.38))
    set_fill(chevron, CYAN)
    set_line(chevron, None)

    items = [
        (1.00, "APR 07", "HBM·헬륨", ["공급부족 5년+", "M15X 생산기반", "헬륨 단기 비축"], CYAN, True, "[S1]"),
        (3.32, "APR 08", "메모리·장비", ["DDR3까지 풍선효과", "CXMT 격차 3년", "테스트 +55%"], BLUE, False, "[S2]"),
        (5.64, "APR 14", "기판·패키징", ["FC-BGA 2조원", "CoPoS 전환", "낸드·eSSD 회복"], LIME, True, "[S3]"),
        (7.96, "APR 22", "증설·원가", ["P5 쌍둥이 팹", "CCL 연쇄 인상", "소재 원가 +20~30%"], ORANGE, False, "[S4]"),
        (10.28, "APR 28", "시스템·효율", ["CPU·CXL로 확장", "전력당 성능", "공급자 우위 고착"], PURPLE, True, "[S5]"),
    ]

    for x, date, title, bullets, c, top, src in items:
        add_circle(slide, x + 0.65, 3.43, 0.38, fill=c, line=BG, width=2)
        if top:
            cy = 1.63
            add_line(slide, x + 0.84, 2.98, x + 0.84, 3.43, c, 1.2)
        else:
            cy = 4.06
            add_line(slide, x + 0.84, 3.81, x + 0.84, 4.06, c, 1.2)
        add_box(slide, x, cy, 2.02, 1.48, fill=PANEL, line=GRID)
        add_pill(slide, x + 0.16, cy + 0.14, 0.76, 0.25, date, fill=c, color=BG, size=8.4)
        add_text(slide, x + 0.16, cy + 0.49, 1.66, 0.25, title, size=12.5, color=WHITE, bold=True)
        add_multiline(slide, x + 0.16, cy + 0.79, 1.65, 0.53, [
            {"text": f"• {b}", "size": 8.7, "color": MUTED, "before": 1} for b in bullets
        ])
        add_text(slide, x + 1.58, cy + 0.15, 0.25, 0.16, src, size=6.8, color=DIM, align=PP_ALIGN.RIGHT)

    add_box(slide, 0.95, 6.00, 11.53, 0.64, fill=PANEL_2, line=None)
    add_text(slide, 1.20, 6.11, 10.98, 0.34, "수요 충격 → 메모리 부족 → 패키징·기판 증설 → 소재·전력 압박 → 시스템 최적화 경쟁", size=15.3, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_footer(slide, 3)
    return slide


def slide_structural_shift(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 4, "STRUCTURAL SHIFT", "경기민감 제조업에서 ‘AI 인프라 계약산업’으로", "수요·계약·경쟁·협상력의 네 축이 동시에 전환")

    # Before panel
    add_box(slide, 0.65, 1.56, 5.15, 3.72, fill=PANEL, line=GRID)
    add_pill(slide, 0.92, 1.82, 1.15, 0.31, "BEFORE", fill=PANEL_3, color=MUTED, size=10)
    before = [
        ("수요", "PC·모바일 교체 사이클", "01"),
        ("계약", "분기·현물 중심 조달", "02"),
        ("경쟁", "미세공정·단품 성능", "03"),
        ("권력", "대형 고객 중심 구매자 우위", "04"),
    ]
    for i, (k, v, n) in enumerate(before):
        yy = 2.30 + i * 0.68
        add_number_badge(slide, 0.95, yy, n, DIM, 0.34)
        add_text(slide, 1.43, yy - 0.01, 0.70, 0.22, k, size=10, color=DIM, bold=True)
        add_text(slide, 2.16, yy - 0.04, 3.20, 0.28, v, size=12.5, color=INK, bold=True)
        if i < 3:
            add_line(slide, 1.43, yy + 0.42, 5.42, yy + 0.42, GRID, 0.6)

    # Center transition arrow
    arrow = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(5.98), Inches(2.65), Inches(1.34), Inches(1.40))
    set_fill(arrow, CYAN)
    set_line(arrow, None)
    add_text(slide, 6.13, 3.00, 0.80, 0.50, "AI", size=21, color=BG, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_text(slide, 5.82, 4.27, 1.55, 0.42, "INFRASTRUCTURE\nECONOMY", size=8.7, color=CYAN, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)

    # Now panel
    add_box(slide, 7.53, 1.56, 5.15, 3.72, fill=PANEL_2, line=CYAN, width=1.2)
    add_pill(slide, 7.80, 1.82, 1.15, 0.31, "NOW", fill=CYAN, color=BG, size=10)
    now = [
        ("수요", "AI 학습·추론의 구조적 수요", "01", CYAN),
        ("계약", "선수금·3~5년 LTA", "02", BLUE),
        ("경쟁", "메모리+패키징+SW 통합", "03", LIME),
        ("권력", "병목 공급자의 가격결정력", "04", ORANGE),
    ]
    for i, (k, v, n, c) in enumerate(now):
        yy = 2.30 + i * 0.68
        add_number_badge(slide, 7.83, yy, n, c, 0.34)
        add_text(slide, 8.31, yy - 0.01, 0.70, 0.22, k, size=10, color=c, bold=True)
        add_text(slide, 9.05, yy - 0.04, 3.20, 0.28, v, size=12.5, color=WHITE, bold=True)
        if i < 3:
            add_line(slide, 8.31, yy + 0.42, 12.30, yy + 0.42, GRID, 0.6)

    # Implications bottom
    cards = [
        ("REVENUE", "가격보다 ‘용량 예약’이 먼저", CYAN),
        ("CAPEX", "고객 커버리지 연동 집행", BLUE),
        ("TECH", "패키징·전성비가 차별화", LIME),
        ("RISK", "소재·지역·IP 동시 헤지", ORANGE),
    ]
    for i, (tag, body, c) in enumerate(cards):
        x = 0.65 + i * 3.01
        add_box(slide, x, 5.61, 2.86, 0.91, fill=PANEL, line=GRID)
        add_text(slide, x + 0.18, 5.78, 0.74, 0.18, tag, size=8, color=c, bold=True, font=FONT_EN)
        add_text(slide, x + 0.18, 6.04, 2.48, 0.25, body, size=10.4, color=INK, bold=True)
    add_footer(slide, 4)
    return slide


def slide_bottleneck_chain(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 5, "BOTTLENECK MAP", "병목은 GPU에서 끝나지 않는다 — 오른쪽으로 이동하며 가치가 확산", "메모리 → 패키징 → 기판·테스트 → 소재·에너지 순으로 다음 제약이 등장")

    stages = [
        ("01", "COMPUTE", "GPU · ASIC", "AI 투자 출발점", "칩 확보", BLUE),
        ("02", "MEMORY", "HBM · DRAM · eSSD", "5년+ 공급부족 전망", "LTA/가격", CYAN),
        ("03", "PACKAGE", "CoWoS · CoPoS · Bonding", "수율·리드타임 제약", "생산속도", PURPLE),
        ("04", "SUBSTRATE / TEST", "FC-BGA · Probe", "주문이 CAPA 상회", "품질/수율", LIME),
        ("05", "MATERIAL / POWER", "He · CCL · 화학 · 전력", "원가·지역집중 노출", "복원력", ORANGE),
    ]
    x_positions = [0.65, 3.08, 5.51, 7.94, 10.37]
    widths = [2.18] * 5

    # top directional guide
    add_text(slide, 0.68, 1.52, 2.0, 0.20, "BOTTLENECK MIGRATION", size=8.5, color=DIM, bold=True, font=FONT_EN)
    add_line(slide, 2.45, 1.64, 12.10, 1.64, GRID, 2.6)
    chev = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(11.92), Inches(1.48), Inches(0.46), Inches(0.31))
    set_fill(chev, ORANGE)
    set_line(chev, None)

    for i, (num, en, items, signal, key, c) in enumerate(stages):
        x = x_positions[i]
        add_box(slide, x, 2.00, widths[i], 3.20, fill=PANEL if i < 4 else PANEL_2, line=c if i == 4 else GRID, width=1.3 if i == 4 else 0.8)
        add_number_badge(slide, x + 0.17, 2.18, num, c, 0.38)
        add_text(slide, x + 0.64, 2.22, 1.33, 0.18, en, size=8.4, color=c, bold=True, font=FONT_EN)
        # simple native icon: nested chip squares / stack lines
        if i in [0, 1]:
            add_box(slide, x + 0.62, 2.72, 0.94, 0.66, fill=PANEL_3, line=c, radius=False, width=1.3)
            if i == 1:
                for k in range(3):
                    add_line(slide, x + 0.72, 2.84 + 0.14 * k, x + 1.46, 2.84 + 0.14 * k, c, 1)
            else:
                add_box(slide, x + 0.86, 2.87, 0.45, 0.34, fill=c, line=None, radius=False)
        elif i == 2:
            for k in range(3):
                add_box(slide, x + 0.57 + 0.14 * k, 2.79 - 0.06 * k, 0.90, 0.52, fill=PANEL_3, line=c, radius=False, width=1)
        elif i == 3:
            add_box(slide, x + 0.53, 2.82, 1.10, 0.49, fill=PANEL_3, line=c, radius=False)
            for k in range(5):
                add_circle(slide, x + 0.58 + 0.20 * k, 2.96, 0.08, fill=c, line=None)
        else:
            add_circle(slide, x + 0.73, 2.72, 0.72, fill=None, line=c, width=1.5)
            add_text(slide, x + 0.73, 2.84, 0.72, 0.30, "He", size=15, color=c, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)

        add_text(slide, x + 0.17, 3.60, 1.84, 0.48, items, size=11.0, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        add_line(slide, x + 0.25, 4.18, x + 1.93, 4.18, GRID, 0.7)
        add_text(slide, x + 0.18, 4.33, 1.82, 0.32, signal, size=9.3, color=MUTED, align=PP_ALIGN.CENTER)
        add_pill(slide, x + 0.53, 4.77, 1.12, 0.28, key, fill=PANEL_3, color=c, size=8.7)
        if i < 4:
            arrow = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x + 2.21), Inches(3.23), Inches(0.44), Inches(0.56))
            set_fill(arrow, GRID)
            set_line(arrow, None)

    # Opportunity strip
    add_box(slide, 0.65, 5.53, 12.03, 1.03, fill=PANEL_2, line=None)
    add_pill(slide, 0.89, 5.74, 1.50, 0.30, "VALUE MIGRATION", fill=CYAN, color=BG, size=9)
    add_text(slide, 2.64, 5.67, 9.65, 0.31, "수혜 범위가 메모리 대기업에서 장비·기판·테스트·소재 기업으로 확산", size=15.6, color=WHITE, bold=True)
    add_text(slide, 2.64, 6.08, 9.50, 0.25, "동시에 마지막 병목(소재·전력)은 전체 밸류체인의 생산중단 리스크를 증폭", size=10.4, color=ORANGE, bold=True)
    add_footer(slide, 5)
    return slide


def slide_money_flow(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 6, "CAPITAL FLOW", "돈은 ‘미세화’보다 병목 해소 역량으로 더 빠르게 이동", "테스트·패키징 성장률이 전체 장비 시장을 크게 상회")

    # Left chart panel
    add_box(slide, 0.65, 1.48, 6.05, 4.80, fill=PANEL, line=GRID)
    add_text(slide, 0.95, 1.77, 4.1, 0.28, "2025 글로벌 반도체 장비 시장", size=13, color=WHITE, bold=True)
    add_rich_text(slide, 0.95, 2.12, 4.9, 0.40, [
        {"text": "$135.1B", "size": 25, "color": CYAN, "bold": True, "font": FONT_EN},
        {"text": "  ·  사상 최고권", "size": 11, "color": MUTED, "bold": True},
    ])
    add_text(slide, 0.95, 2.54, 4.9, 0.20, "매출 증가율 (YoY)", size=9, color=DIM, bold=True)

    bars = [
        ("전체 장비", 15, 1.05, BLUE),
        ("조립·패키징", 21, 1.48, PURPLE),
        ("테스트", 55, 3.15, LIME),
    ]
    base_y = 5.55
    for i, (label, val, bh, c) in enumerate(bars):
        bx = 1.18 + i * 1.62
        add_box(slide, bx, base_y - bh, 0.82, bh, fill=c, line=None, radius=False)
        add_text(slide, bx - 0.12, base_y - bh - 0.36, 1.06, 0.28, f"+{val}%", size=14, color=c, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
        add_text(slide, bx - 0.27, 5.70, 1.36, 0.38, label, size=9.1, color=INK, bold=True, align=PP_ALIGN.CENTER)
    # y-axis guides
    for j, pct in enumerate([0, 20, 40, 60]):
        yy = 5.55 - (pct / 60) * 3.44
        add_line(slide, 0.96, yy, 6.34, yy, GRID, 0.55)
        add_text(slide, 0.75, yy - 0.09, 0.18, 0.15, str(pct), size=6.8, color=DIM, align=PP_ALIGN.RIGHT)
    add_pill(slide, 4.96, 1.77, 1.40, 0.28, "[S2] SEMI", fill=PANEL_3, color=BLUE, size=8.2)

    # Right investment panels
    add_text(slide, 7.03, 1.56, 5.55, 0.25, "CAPACITY SIGNALS", size=9, color=DIM, bold=True, font=FONT_EN)
    inv = [
        ("2조원", "FC-BGA", "삼성전기 베트남 증설", "AI 기판 수요 > 기존 CAPA", CYAN, "[S3]"),
        ("160조원", "MEMORY FAB", "평택 P5 쌍둥이 팹 구상", "HBM 생산속도 극대화", ORANGE, "[S4]"),
        ("5,000호", "TEST SCALE", "세메스 프로브 스테이션", "HBM 검사 인프라 고도화", BLUE, "[S2]"),
    ]
    for i, (big, tag, title, note, c, src) in enumerate(inv):
        yy = 1.91 + i * 1.35
        add_box(slide, 7.03, yy, 5.65, 1.14, fill=PANEL, line=GRID)
        add_line(slide, 7.25, yy + 0.22, 7.25, yy + 0.92, c, 3.5)
        add_text(slide, 7.50, yy + 0.17, 1.55, 0.34, big, size=20, color=c, bold=True, font=FONT_EN)
        add_pill(slide, 9.05, yy + 0.18, 1.12, 0.24, tag, fill=PANEL_3, color=c, size=7.5)
        add_text(slide, 7.50, yy + 0.58, 2.60, 0.26, title, size=10.6, color=WHITE, bold=True)
        add_text(slide, 10.34, yy + 0.52, 1.98, 0.33, note, size=9.2, color=MUTED, align=PP_ALIGN.RIGHT)
        add_text(slide, 12.28, yy + 0.90, 0.19, 0.12, src, size=6.5, color=DIM, align=PP_ALIGN.RIGHT)

    add_box(slide, 7.03, 6.02, 5.65, 0.48, fill=PANEL_3, line=None)
    add_text(slide, 7.30, 6.11, 5.13, 0.26, "투자 기준: 고객 커버리지 × 수율 × 병목 해소 효과", size=12.2, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_footer(slide, 6, "자료: [S2] Apr 08, [S3] Apr 14, [S4] Apr 22 · 기사상 계획/전망 기준")
    return slide


def slide_competition(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 7, "COMPETITIVE MAP", "경쟁의 단위가 ‘기업’에서 ‘생태계’로 커졌다", "한국은 메모리, TSMC는 플랫폼, 중국은 속도와 규모로 서로 다른 우위를 구축")

    cols = [
        (0.65, "KOREA", "메모리 + 턴키", CYAN,
         ["HBM·DRAM 리더십", "대규모 팹·패키징 투자", "메모리–파운드리 결합"],
         ["중국 추격", "소재 수입 집중"]),
        (4.72, "TSMC", "파운드리 플랫폼", BLUE,
         ["2nm + CoWoS/CoPoS", "설계·패키징 고객 락인", "글로벌 생산 분산"],
         ["전력·CAPA", "대만 집중 리스크"]),
        (8.79, "CHINA", "자립 + 속도", ORANGE,
         ["CXMT 12단 HBM 추진", "장비·AI칩 내재화", "정부자본·내수 규모"],
         ["수율·장비 제약", "수출통제"]),
    ]
    for x, region, thesis, c, strengths, risks in cols:
        add_box(slide, x, 1.52, 3.89, 3.78, fill=PANEL, line=c, width=1.15)
        add_pill(slide, x + 0.24, 1.79, 1.10, 0.30, region, fill=c, color=BG, size=9.5)
        add_text(slide, x + 0.24, 2.19, 3.35, 0.34, thesis, size=18, color=WHITE, bold=True)
        add_text(slide, x + 0.24, 2.67, 1.1, 0.17, "STRATEGIC EDGE", size=7.8, color=c, bold=True, font=FONT_EN)
        add_multiline(slide, x + 0.24, 2.96, 3.25, 0.98, [
            {"text": f"• {s}", "size": 10.4, "color": INK, "before": 2} for s in strengths
        ])
        add_line(slide, x + 0.24, 4.10, x + 3.59, 4.10, GRID, 0.7)
        add_text(slide, x + 0.24, 4.28, 0.92, 0.17, "WATCH-OUT", size=7.8, color=RED, bold=True, font=FONT_EN)
        add_text(slide, x + 1.16, 4.23, 2.35, 0.48, " · ".join(risks), size=9.8, color=MUTED, bold=True)
        # mini edge indicator
        add_line(slide, x + 0.24, 4.88, x + 3.59, 4.88, GRID, 4)
        for j in range(4):
            add_circle(slide, x + 0.38 + j * 0.80, 4.79, 0.18, fill=c if j < (4 if region == "KOREA" else 3) else GRID, line=None)

    # Battleground strip
    add_box(slide, 0.65, 5.57, 12.03, 0.94, fill=PANEL_2, line=None)
    add_text(slide, 0.90, 5.78, 1.15, 0.19, "BATTLEGROUNDS", size=8, color=DIM, bold=True, font=FONT_EN)
    fields = [
        ("AI MEMORY", "한국 우위", CYAN),
        ("FOUNDRY ECOSYSTEM", "TSMC 우위", BLUE),
        ("COST / SCALE", "중국 상승", ORANGE),
        ("SUPPLY ACCESS", "지정학 변수", PURPLE),
    ]
    for i, (f, state, c) in enumerate(fields):
        x = 2.23 + i * 2.52
        add_circle(slide, x, 5.78, 0.22, fill=c, line=None)
        add_text(slide, x + 0.34, 5.73, 1.95, 0.17, f, size=7.6, color=DIM, bold=True, font=FONT_EN)
        add_text(slide, x + 0.34, 6.02, 1.95, 0.23, state, size=10.6, color=WHITE, bold=True)
    add_footer(slide, 7)
    return slide


def slide_tech_roadmap(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 8, "TECH ROADMAP", "2026–2029: 선폭보다 ‘전력·패키징·수율’이 양산 속도를 좌우", "서로 다른 기술 로드맵이 하나의 AI 시스템 성능으로 결합")

    left = 2.03
    cols = [left, 4.58, 7.13, 9.68]
    years = [("2026", CYAN), ("2027", BLUE), ("2028", PURPLE), ("2029+", ORANGE)]
    for i, (yr, c) in enumerate(years):
        add_pill(slide, cols[i], 1.54, 2.20, 0.33, yr, fill=c, color=BG, size=10.5)
        if i < 3:
            arrow = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(cols[i] + 2.25), Inches(1.59), Inches(0.28), Inches(0.23))
            set_fill(arrow, GRID)
            set_line(arrow, None)

    rows = [
        ("MEMORY", 2.15, CYAN),
        ("LOGIC", 3.48, BLUE),
        ("PACKAGE / SYSTEM", 4.81, PURPLE),
    ]
    for label, y, c in rows:
        add_pill(slide, 0.65, y + 0.34, 1.12, 0.30, label, fill=PANEL_3, color=c, size=8.1)
        add_line(slide, 1.85, y + 0.49, 12.10, y + 0.49, GRID, 1.0)

    milestones = {
        (0, 0): ("HBM4 양산 경쟁", "대역폭·발열·수율", CYAN),
        (0, 1): ("12단 HBM 추격", "CXMT 양산 계획", ORANGE),
        (0, 2): ("CUSTOM HBM", "고객별 로직 결합", BLUE),
        (0, 3): ("차세대 적층", "하이브리드 본딩 경제성", PURPLE),
        (1, 0): ("2nm GAA", "선택적 증착·전성비", BLUE),
        (1, 1): ("수율 램프업", "고객 인증이 분기점", CYAN),
        (1, 2): ("1.4nm 준비", "설계 생태계 선점", PURPLE),
        (1, 3): ("1.4nm 목표", "TSMC 2028 / 삼성 2029", ORANGE),
        (2, 0): ("CoWoS 병목", "FC-BGA·테스트 확대", LIME),
        (2, 1): ("Vertical Die", "I/O 10× · BW 4×", CYAN),
        (2, 2): ("CoPoS 양산", "패널 기반 생산성", PURPLE),
        (2, 3): ("CXL / 통합", "CPU·메모리 공유 구조", BLUE),
    }
    row_y = [2.15, 3.48, 4.81]
    for (r, cidx), (ttl, sub, c) in milestones.items():
        x = cols[cidx]
        y = row_y[r]
        add_box(slide, x, y, 2.20, 1.00, fill=PANEL, line=GRID)
        add_line(slide, x + 0.14, y + 0.18, x + 0.14, y + 0.82, c, 3)
        add_text(slide, x + 0.30, y + 0.17, 1.72, 0.25, ttl, size=10.6, color=WHITE, bold=True)
        add_text(slide, x + 0.30, y + 0.51, 1.72, 0.27, sub, size=8.5, color=MUTED, bold=True)

    add_box(slide, 0.65, 6.16, 12.03, 0.46, fill=PANEL_3, line=None)
    add_text(slide, 0.90, 6.24, 11.55, 0.26, "공통 GATE  |  전력당 성능  ·  열 관리  ·  패키징 수율  ·  고객 인증  ·  소재 안정성", size=11.8, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_footer(slide, 8, "자료: [S2] Apr 08, [S3] Apr 14, [S4] Apr 22, [S5] Apr 28 · 기사상 일정/목표 기준")
    return slide


def slide_risk_heatmap(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 9, "RISK RADAR", "상승장의 최대 리스크는 ‘수요’보다 공급망·실행력에 있다", "확률과 충격을 분리해 관리하고 선행지표로 조기경보")

    # Heatmap panel
    x0, y0, pw, ph = 0.90, 1.78, 7.35, 4.62
    # quadrants
    add_box(slide, x0, y0, pw / 2, ph / 2, fill="0C2230", line=None, radius=False)
    add_box(slide, x0 + pw / 2, y0, pw / 2, ph / 2, fill="2B2224", line=None, radius=False)
    add_box(slide, x0, y0 + ph / 2, pw / 2, ph / 2, fill="0A1D2A", line=None, radius=False)
    add_box(slide, x0 + pw / 2, y0 + ph / 2, pw / 2, ph / 2, fill="182534", line=None, radius=False)
    add_line(slide, x0 + pw / 2, y0, x0 + pw / 2, y0 + ph, GRID, 1.0)
    add_line(slide, x0, y0 + ph / 2, x0 + pw, y0 + ph / 2, GRID, 1.0)
    add_text(slide, 0.66, 3.50, 0.23, 1.10, "영향도", size=9, color=DIM, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_text(slide, 3.42, 6.52, 2.30, 0.22, "발생 가능성  →", size=9, color=DIM, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, x0 + 0.18, y0 + 0.14, 1.10, 0.18, "HIGH IMPACT", size=7.5, color=MUTED, bold=True, font=FONT_EN)
    add_text(slide, x0 + pw - 1.46, y0 + 0.14, 1.20, 0.18, "ACT NOW", size=8, color=RED, bold=True, font=FONT_EN, align=PP_ALIGN.RIGHT)

    risks = [
        (5.65, 2.36, 0.88, "소재·헬륨\n공급차질", ORANGE),
        (6.54, 3.12, 0.90, "패키징\nCAPA", RED),
        (4.55, 2.88, 0.82, "중국\n추격", PURPLE),
        (3.25, 2.33, 0.88, "CAPEX\n과잉", BLUE),
        (5.72, 4.45, 0.80, "CCL·화학\n원가", YELLOW),
        (3.89, 4.75, 0.72, "인재·IP\n유출", CYAN),
        (2.05, 3.95, 0.80, "수율\n램프업", LIME),
    ]
    for x, y, d, label, c in risks:
        add_circle(slide, x, y, d, fill=c, line=None)
        add_text(slide, x + 0.05, y + 0.10, d - 0.10, d - 0.17, label, size=8.4, color=BG, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    # Right indicators
    add_box(slide, 8.60, 1.78, 4.08, 4.62, fill=PANEL, line=GRID)
    add_pill(slide, 8.90, 2.05, 1.54, 0.29, "EARLY WARNINGS", fill=PANEL_3, color=CYAN, size=8.5)
    indicators = [
        ("01", "HBM·패키징 리드타임", "주간", CYAN),
        ("02", "핵심소재 재고 커버리지", "주간", ORANGE),
        ("03", "2nm/HBM 수율·고객 인증", "월간", LIME),
        ("04", "CCL·나프타·화학 가격", "월간", YELLOW),
        ("05", "CXMT·YMTC 고객 채택", "분기", PURPLE),
        ("06", "LTA·선수금 커버리지", "분기", BLUE),
    ]
    for i, (num, label, freq, c) in enumerate(indicators):
        yy = 2.60 + i * 0.56
        add_number_badge(slide, 8.92, yy, num, c, 0.30)
        add_text(slide, 9.38, yy - 0.01, 2.55, 0.20, label, size=9.7, color=INK, bold=True)
        add_pill(slide, 11.92, yy - 0.01, 0.50, 0.22, freq, fill=PANEL_3, color=MUTED, size=6.9)
        if i < 5:
            add_line(slide, 9.38, yy + 0.37, 12.38, yy + 0.37, GRID, 0.5)

    add_box(slide, 8.88, 6.02, 3.52, 0.18, fill=RED, line=None, radius=False)
    add_text(slide, 8.90, 5.69, 3.50, 0.24, "RED FLAG", size=8, color=RED, bold=True, font=FONT_EN)
    add_text(slide, 10.12, 5.64, 2.28, 0.31, "2개 이상 동시 악화 시\n투자 Gate 재심의", size=9.5, color=WHITE, bold=True, align=PP_ALIGN.RIGHT)
    add_footer(slide, 9)
    return slide


def slide_opportunity_portfolio(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 10, "VALUE POOLS", "기회는 크지만 자본강도와 진입속도가 다르다", "버블 크기 = 상대적 자본강도 · 위치 = 시장 매력도와 진입 난도(정성 평가)")

    x0, y0, pw, ph = 0.92, 1.74, 8.22, 4.83
    # portfolio area
    add_box(slide, x0, y0, pw, ph, fill=PANEL, line=GRID, radius=False)
    for frac in [0.33, 0.66]:
        add_line(slide, x0 + pw * frac, y0, x0 + pw * frac, y0 + ph, GRID, 0.65)
        add_line(slide, x0, y0 + ph * frac, x0 + pw, y0 + ph * frac, GRID, 0.65)
    add_text(slide, 0.63, 3.34, 0.22, 1.15, "시장 매력도", size=9, color=DIM, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_text(slide, 3.57, 6.67, 2.90, 0.20, "진입 난도 / CAPEX  →", size=9, color=DIM, bold=True, align=PP_ALIGN.CENTER)
    add_pill(slide, 1.12, 1.94, 1.38, 0.25, "QUICK WINS", fill="12312F", color=CYAN, size=7.8)
    add_pill(slide, 7.39, 1.94, 1.38, 0.25, "BIG BETS", fill="2B2434", color=PURPLE, size=7.8)

    bubbles = [
        (2.06, 2.50, 1.02, "테스트·\n검사", LIME),
        (3.37, 2.18, 1.22, "AI 기판·\nFC-BGA", CYAN),
        (5.14, 2.58, 1.18, "첨단\n패키징", PURPLE),
        (7.00, 2.05, 1.55, "HBM·\nCUSTOM", BLUE),
        (2.10, 4.15, 0.95, "특수가스·\n소재", ORANGE),
        (4.15, 4.08, 1.10, "eSSD·\nNAND", CYAN),
        (6.05, 4.48, 1.16, "K-NPU·\nEDGE", LIME),
        (7.67, 4.60, 1.32, "2nm\n통합", RED),
        (3.10, 5.38, 0.79, "레거시\n메모리", YELLOW),
    ]
    for x, y, d, label, c in bubbles:
        add_circle(slide, x, y, d, fill=c, line=None)
        add_text(slide, x + 0.05, y + 0.08, d - 0.10, d - 0.15, label, size=8.4 if d < 1.0 else 9.3, color=BG, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    # Right allocation logic
    add_box(slide, 9.48, 1.74, 3.20, 4.83, fill=PANEL_2, line=GRID)
    add_pill(slide, 9.76, 2.00, 1.42, 0.29, "ALLOCATION LOGIC", fill=CYAN, color=BG, size=8.2)
    logic = [
        ("1", "Protect", "공급중단을 막는 소재·테스트", ORANGE),
        ("2", "Scale", "고객수요가 확인된 기판·패키징", CYAN),
        ("3", "Partner", "CAPEX 큰 HBM·2nm는 공동투자", BLUE),
        ("4", "Option", "K-NPU·CXL은 파일럿으로 학습", LIME),
    ]
    for i, (n, en, ko, c) in enumerate(logic):
        yy = 2.66 + i * 0.82
        add_number_badge(slide, 9.78, yy, n, c, 0.38)
        add_text(slide, 10.35, yy - 0.01, 1.98, 0.18, en.upper(), size=8, color=c, bold=True, font=FONT_EN)
        add_text(slide, 10.35, yy + 0.26, 1.98, 0.36, ko, size=9.6, color=INK, bold=True)
    add_line(slide, 9.76, 5.98, 12.38, 5.98, GRID, 0.8)
    add_text(slide, 9.76, 6.10, 2.62, 0.28, "당사 역량·고객 포트폴리오에 맞춰\n버블 위치를 재평가해야 함", size=8.7, color=MUTED, align=PP_ALIGN.CENTER)
    add_footer(slide, 10, "자료: [S1–S5] 기반 정성적 전략 평가 · 당사 역량 반영 전 가설")
    return slide


def slide_scenarios(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 11, "SCENARIO", "2026–2027: 방향성은 성장, 차이는 ‘지속기간과 병목 강도’", "단일 전망보다 트리거 기반으로 투자·가격·재고를 조정")

    scenarios = [
        (0.65, "UPSIDE", "AI CAPEX 재가속", LIME,
         ["LTA·선수금 확대", "패키징 리드타임 재상승", "HBM4/ASIC 동시 강세"],
         "용량 선점 · 가격 옵션 확보"),
        (4.72, "BASE", "타이트한 정상화", CYAN,
         ["공급부족 지속", "가격 상승폭 완만", "병목이 소재·전력으로 이동"],
         "고객연계 증설 · 믹스 개선"),
        (8.79, "DOWNSIDE", "AI 수익화 지연", RED,
         ["빅테크 CAPEX 연기", "레거시 가격 조정", "중국 공급 확대"],
         "Gate 중단 · 현금/재고 방어"),
    ]
    for x, name, title, c, triggers, action in scenarios:
        add_box(slide, x, 1.56, 3.89, 4.68, fill=PANEL, line=c, width=1.2)
        add_pill(slide, x + 0.24, 1.83, 1.03, 0.30, name, fill=c, color=BG, size=9)
        add_text(slide, x + 0.24, 2.24, 3.34, 0.35, title, size=18, color=WHITE, bold=True)
        add_text(slide, x + 0.24, 2.79, 0.85, 0.18, "TRIGGERS", size=7.8, color=DIM, bold=True, font=FONT_EN)
        for i, t in enumerate(triggers):
            yy = 3.11 + i * 0.52
            add_circle(slide, x + 0.27, yy + 0.04, 0.13, fill=c, line=None)
            add_text(slide, x + 0.54, yy, 2.93, 0.25, t, size=10.2, color=INK, bold=True)
        add_line(slide, x + 0.24, 4.77, x + 3.60, 4.77, GRID, 0.8)
        add_text(slide, x + 0.24, 4.96, 0.90, 0.18, "PLAYBOOK", size=7.8, color=c, bold=True, font=FONT_EN)
        add_box(slide, x + 0.24, 5.28, 3.35, 0.55, fill=PANEL_3, line=None)
        add_text(slide, x + 0.38, 5.38, 3.07, 0.30, action, size=11.3, color=WHITE, bold=True, align=PP_ALIGN.CENTER)

    add_box(slide, 0.65, 6.48, 12.03, 0.34, fill=CYAN, line=None, radius=False)
    add_text(slide, 0.88, 6.53, 11.56, 0.20, "NO-REGRET MOVE  |  LTA 확대 · Dual source · 수율 Gate · 모듈형 CAPEX · 핵심인재/기술 보호", size=10.8, color=BG, bold=True, align=PP_ALIGN.CENTER)
    add_footer(slide, 11)
    return slide


def slide_ceo_decisions(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 12, "CEO DECISIONS", "지금 승인해야 할 5가지 선택", "‘성장’과 ‘방어’를 동시에 달성하는 의사결정 패키지")

    decisions = [
        ("01", "포트폴리오", "AI/HPC 연계 고부가 영역으로 자원 이동", "매출·GM 믹스", CYAN, "D+30"),
        ("02", "고객계약", "LTA·용량예약·선수금 옵션 표준화", "수요 커버리지", BLUE, "D+45"),
        ("03", "CAPEX", "고객·수율·병목 3-Gate 충족분만 집행", "ROIC / 수율", LIME, "D+30"),
        ("04", "공급망", "He·CCL·화학 핵심품목 이원화·비축 기준 승인", "재고 커버리지", ORANGE, "D+60"),
        ("05", "기술·인재", "패키징·테스트·전성비 핵심인력/IP 락인", "핵심인재 유지", PURPLE, "D+90"),
    ]
    for i, (num, area, action, kpi, c, due) in enumerate(decisions):
        y = 1.56 + i * 0.99
        add_box(slide, 0.65, y, 12.03, 0.82, fill=PANEL, line=GRID)
        add_number_badge(slide, 0.91, y + 0.21, num, c, 0.40)
        add_text(slide, 1.51, y + 0.17, 1.18, 0.22, area, size=11.3, color=c, bold=True)
        add_text(slide, 2.76, y + 0.16, 5.50, 0.27, action, size=13.4, color=WHITE, bold=True)
        add_pill(slide, 8.54, y + 0.20, 1.92, 0.28, f"KPI  {kpi}", fill=PANEL_3, color=MUTED, size=8)
        add_pill(slide, 10.70, y + 0.20, 0.82, 0.28, due, fill=c, color=BG, size=8.4)
        add_pill(slide, 11.73, y + 0.20, 0.66, 0.28, "승인", fill=PANEL_3, color=c, size=8.4, line=c)

    add_box(slide, 0.65, 6.62, 12.03, 0.25, fill=PANEL_3, line=None, radius=False)
    add_text(slide, 0.88, 6.64, 11.56, 0.18, "의사결정 원칙  |  용량보다 고객 · 매출보다 현금 · 기술보다 양산 · 단일소싱보다 복원력", size=9.6, color=INK, bold=True, align=PP_ALIGN.CENTER)
    add_footer(slide, 12)
    return slide


def slide_90day(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 13, "90-DAY PLAN", "결정을 실행으로 바꾸는 90일", "진단 → 계약·공급 잠금 → 투자·운영 전환의 3단계")

    phase_x = [2.32, 5.70, 9.08]
    phases = [
        ("0–30 DAYS", "DIAGNOSE", CYAN),
        ("31–60 DAYS", "LOCK-IN", BLUE),
        ("61–90 DAYS", "EXECUTE", LIME),
    ]
    for i, (days, name, c) in enumerate(phases):
        add_pill(slide, phase_x[i], 1.48, 3.08, 0.31, days, fill=c, color=BG, size=9.2)
        add_text(slide, phase_x[i], 1.86, 3.08, 0.23, name, size=9, color=c, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
        if i < 2:
            arrow = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(phase_x[i] + 3.13), Inches(1.52), Inches(0.31), Inches(0.23))
            set_fill(arrow, GRID)
            set_line(arrow, None)

    rows = [
        ("COMMERCIAL", CYAN), ("SUPPLY", ORANGE), ("TECH / CAPEX", PURPLE), ("RISK / ORG", RED)
    ]
    content = [
        ["Top 고객 수요·LTA 갭 진단", "용량예약·가격 옵션 제안", "계약 전환·믹스 실행"],
        ["He·CCL·화학 취약점 맵", "2nd source·비축 기준 합의", "조달계약·재고 재배치"],
        ["병목·수율·CAPEX 재평가", "3-Gate 투자안 확정", "파일럿/증설 집행"],
        ["중국·IP·인재 위험 점검", "핵심인력/기술 보호안", "CEO 대시보드 가동"],
    ]
    for r, (label, c) in enumerate(rows):
        y = 2.32 + r * 0.90
        add_pill(slide, 0.65, y + 0.20, 1.35, 0.30, label, fill=PANEL_3, color=c, size=7.8)
        for col in range(3):
            x = phase_x[col]
            add_box(slide, x, y, 3.08, 0.70, fill=PANEL if (r + col) % 2 == 0 else PANEL_2, line=GRID)
            add_circle(slide, x + 0.20, y + 0.26, 0.16, fill=c, line=None)
            add_text(slide, x + 0.52, y + 0.19, 2.30, 0.31, content[r][col], size=9.9, color=INK, bold=True, valign=MSO_ANCHOR.MIDDLE)

    # KPI strip
    add_box(slide, 0.65, 6.10, 12.03, 0.65, fill=PANEL_3, line=None)
    add_text(slide, 0.88, 6.21, 0.70, 0.18, "WEEKLY KPI", size=7.8, color=CYAN, bold=True, font=FONT_EN)
    kpis = ["LTA 커버리지", "핵심소재 주(週)", "패키징 수율", "CAPEX 고객커버", "AI/HPC GM"]
    for i, k in enumerate(kpis):
        x = 1.78 + i * 2.08
        add_box(slide, x, 6.20, 1.84, 0.30, fill=PANEL, line=GRID)
        add_text(slide, x, 6.24, 1.84, 0.17, k, size=8.5, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_footer(slide, 13)
    return slide


def slide_sources(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 14, "APPENDIX", "자료 범위와 해석 원칙", "수치·일정은 각 세미뉴스에 인용된 기사·기관 전망 기준")

    sources = [
        ("S1", "세미뉴스 Apr 07 2026.pdf", "HBM 공급부족 · 헬륨 · M15X · 데이터센터 장비"),
        ("S2", "세미뉴스 Apr 08 2026.pdf", "레거시 메모리 · CXMT · 장비시장 · Vertical Die"),
        ("S3", "세미뉴스 Apr 14 2026.pdf", "FC-BGA · CoPoS · eSSD · 2nm 증착"),
        ("S4", "세미뉴스 Apr 22 2026.pdf", "P4/P5 · CCL · 소재원가 · 전성비"),
        ("S5", "세미뉴스 Apr 28 2026.pdf", "월말 기술·산업 신호 및 종합 방향성"),
    ]
    add_box(slide, 0.65, 1.53, 7.70, 4.76, fill=PANEL, line=GRID)
    add_pill(slide, 0.92, 1.80, 1.12, 0.30, "SOURCE SET", fill=CYAN, color=BG, size=9)
    for i, (sid, name, topics) in enumerate(sources):
        y = 2.33 + i * 0.71
        add_pill(slide, 0.94, y, 0.53, 0.28, sid, fill=PANEL_3, color=CYAN, size=8)
        add_text(slide, 1.66, y - 0.02, 2.93, 0.22, name, size=10.5, color=WHITE, bold=True)
        add_text(slide, 4.63, y - 0.01, 3.30, 0.26, topics, size=8.7, color=MUTED, align=PP_ALIGN.RIGHT)
        if i < 4:
            add_line(slide, 1.66, y + 0.45, 7.98, y + 0.45, GRID, 0.5)

    # Interpretation principles
    add_box(slide, 8.68, 1.53, 4.00, 4.76, fill=PANEL_2, line=GRID)
    add_pill(slide, 8.96, 1.80, 1.42, 0.30, "READING GUIDE", fill=PANEL_3, color=LIME, size=8.5)
    principles = [
        ("01", "FACT", "기사 내 수치·발표·계획", CYAN),
        ("02", "OUTLOOK", "기관·업계의 전망", BLUE),
        ("03", "SYNTHESIS", "5개 이슈 간 공통 패턴", PURPLE),
        ("04", "IMPLICATION", "일반화한 CEO 의사결정 가설", ORANGE),
    ]
    for i, (num, en, ko, c) in enumerate(principles):
        y = 2.48 + i * 0.70
        add_number_badge(slide, 8.98, y, num, c, 0.36)
        add_text(slide, 9.52, y - 0.01, 0.94, 0.18, en, size=7.8, color=c, bold=True, font=FONT_EN)
        add_text(slide, 10.48, y - 0.02, 1.82, 0.25, ko, size=9.4, color=INK, bold=True)
    add_line(slide, 8.96, 5.43, 12.40, 5.43, GRID, 0.8)
    add_text(slide, 8.96, 5.63, 3.44, 0.40, "주의  |  투자액·양산 일정·시장 전망은\n후속 공시와 최신 데이터로 재검증 필요", size=9.2, color=MUTED, bold=True, align=PP_ALIGN.CENTER)

    add_box(slide, 0.65, 6.55, 12.03, 0.29, fill=CYAN, line=None, radius=False)
    add_text(slide, 0.87, 6.59, 11.58, 0.18, "본 자료의 전략 제안은 회사별 제품·고객·재무 상황을 반영하기 전의 CEO 논의용 가설입니다.", size=9.5, color=BG, bold=True, align=PP_ALIGN.CENTER)
    add_footer(slide, 14, "Source set: user-provided April 2026 semiconductor-news files [S1–S5]")
    return slide


# ---------------------------------------------------------------------------
# Build / metadata / validation
# ---------------------------------------------------------------------------


def build() -> Path:
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)

    props = prs.core_properties
    props.title = "2026년 4월 세미뉴스 종합 CEO 보고"
    props.subject = "AI 반도체 구조적 슈퍼사이클과 CEO 의사결정"
    props.author = "Arena.ai"
    props.keywords = "반도체, HBM, AI, 패키징, CEO, 2026년 4월"
    props.comments = "Native-shape, fully editable CEO briefing deck."

    slide_cover(prs)
    slide_executive_summary(prs)
    slide_timeline(prs)
    slide_structural_shift(prs)
    slide_bottleneck_chain(prs)
    slide_money_flow(prs)
    slide_competition(prs)
    slide_tech_roadmap(prs)
    slide_risk_heatmap(prs)
    slide_opportunity_portfolio(prs)
    slide_scenarios(prs)
    slide_ceo_decisions(prs)
    slide_90day(prs)
    slide_sources(prs)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUT)
    return OUT


if __name__ == "__main__":
    out = build()
    print(out)
