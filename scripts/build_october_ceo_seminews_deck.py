#!/usr/bin/env python3
"""Build the October 2026 semiconductor-news CEO briefing.

Input scope: sixteen attached SEMI News for YOU issues dated Oct 02–30, 2026.
The deck uses the same 16:9 dark executive visual system as the April and
September briefings. All diagrams are native, editable PowerPoint shapes.
"""

from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

from build_ceo_seminews_deck import (
    W, H, BG, BG_ALT, PANEL, PANEL_2, PANEL_3, GRID, WHITE, INK, MUTED, DIM,
    CYAN, BLUE, LIME, ORANGE, YELLOW, RED, PURPLE, FONT_EN,
    set_fill, set_line, add_box, add_circle, add_line, add_text, add_rich_text,
    add_multiline, add_pill, add_number_badge, add_header, add_footer,
    add_slide_bg,
)

OUT = Path("deliverables/2026_10_세미뉴스_CEO_보고.pptx")
SRC_DEFAULT = "자료: SEMI News for YOU 2026년 10월 16개 이슈 [S01–S16]"


def footer(slide, number: int, source: str = SRC_DEFAULT):
    add_footer(slide, number, source)


def cover(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)

    for x in [8.00, 8.55, 9.10, 9.65, 10.20, 10.75, 11.30, 11.85, 12.40]:
        add_line(slide, x, 0.56, x, 6.93, GRID, 0.45)
    for y in [0.72, 1.27, 1.82, 2.37, 2.92, 3.47, 4.02, 4.57, 5.12, 5.67, 6.22, 6.77]:
        add_line(slide, 7.65, y, 12.78, y, GRID, 0.45)

    add_pill(slide, 0.75, 0.62, 2.55, 0.32, "CEO BRIEFING · CONFIDENTIAL", fill=PANEL_2, color=CYAN, size=9.5)
    add_text(slide, 0.75, 1.38, 6.85, 0.58, "2026년 10월 세미뉴스", size=32, color=WHITE, bold=True)
    add_text(slide, 0.75, 2.02, 6.95, 1.23, "AI 호황은 ‘칩 부족’을 넘어\n전력·패키징·자본의 배분전쟁으로", size=28.5, color=WHITE, bold=True)
    add_line(slide, 0.75, 3.49, 2.30, 3.49, CYAN, 4)
    add_text(
        slide, 0.75, 3.76, 6.34, 0.94,
        "16개 이슈·258개 헤드라인을 수요, 메모리, 파운드리,\n첨단패키징, 중국 추격, 인프라 관점으로 통합했습니다.",
        size=14.1, color=MUTED,
    )
    add_text(slide, 0.75, 6.35, 6.40, 0.26, "SEMI NEWS  |  OCT 02–30  ·  16 ISSUES  ·  258 SIGNALS", size=9.8, color=CYAN, bold=True, font=FONT_EN)
    add_text(slide, 0.75, 6.69, 5.8, 0.20, "Prepared for CEO · 2026.10", size=9, color=DIM, font=FONT_EN)

    # Editable AI-factory resource orbit
    cx, cy = 10.25, 3.52
    for d, c, lw in [(4.34, GRID, 1.0), (3.58, PANEL_3, 1.0), (2.76, BLUE, 1.2)]:
        add_circle(slide, cx - d/2, cy - d/2, d, fill=None, line=c, width=lw)
    core = add_box(slide, 9.13, 2.39, 2.24, 2.24, fill=PANEL_2, line=CYAN, radius=True, width=2)
    core.rotation = 45
    inner = add_box(slide, 9.67, 2.93, 1.16, 1.16, fill=CYAN, line=None, radius=False)
    inner.rotation = 45
    add_text(slide, 9.41, 3.06, 1.68, 0.72, "AI\nFACTORY", size=11.2, color=BG, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    nodes = [
        (8.08, 1.30, "CUSTOM", BLUE), (9.50, 0.92, "HBM", LIME),
        (11.50, 1.13, "PACKAGE", PURPLE), (12.13, 2.48, "NETWORK", ORANGE),
        (12.06, 4.55, "POWER", YELLOW), (10.93, 5.77, "COOLING", CYAN),
        (8.78, 5.79, "CAPITAL", BLUE), (7.63, 4.34, "SUPPLY", RED),
    ]
    for nx, ny, label, c in nodes:
        add_circle(slide, nx, ny, 0.36, fill=c, line=None)
        add_text(slide, nx-0.23, ny+0.43, 0.82, 0.18, label, size=6.3, color=c, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
        add_line(slide, nx+0.18, ny+0.18, cx, cy, c, 0.8)
    add_pill(slide, 8.07, 6.31, 4.65, 0.32, "CHIP SHORTAGE → RESOURCE ALLOCATION", fill=PANEL_2, color=BLUE, size=8.6)
    return slide


def executive(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 2, "EXECUTIVE", "10월의 결론: 수요 검증은 끝났고, 2028–30 자원 확보전이 시작됐다")

    add_box(slide, 0.65, 1.48, 7.65, 2.17, fill=PANEL, line=GRID)
    add_pill(slide, 0.92, 1.74, 1.18, 0.30, "ONE-LINE", fill=PANEL_3, color=LIME, size=9)
    add_rich_text(slide, 0.92, 2.13, 7.05, 0.91, [
        {"text": "AI 호황의 병목이 ", "size": 22.3, "color": WHITE, "bold": True},
        {"text": "GPU → 메모리 → 패키징 → 전력·자본", "size": 22.3, "color": CYAN, "bold": True},
        {"text": "으로 이동한다.", "size": 22.3, "color": WHITE, "bold": True},
    ])
    add_text(slide, 0.92, 3.08, 6.96, 0.31, "장기계약과 현지 생태계가 가격보다 강한 공급권이 되며, 중국은 규제를 국산화 속도로 전환하고 있습니다.", size=10.9, color=MUTED)

    add_box(slide, 8.55, 1.48, 4.13, 2.17, fill=PANEL_2, line=CYAN, width=1.3)
    add_pill(slide, 8.84, 1.74, 1.53, 0.30, "CEO IMPERATIVE", fill=CYAN, color=BG, size=9)
    add_multiline(slide, 8.84, 2.14, 3.48, 1.18, [
        {"text": "‘증설’이 아니라 공급권을 확보", "size": 17.2, "color": WHITE, "bold": True},
        {"text": "① LTA·공동설계로 수요 선점", "size": 11.1, "color": INK, "before": 8},
        {"text": "② 웨이퍼–패키지–전력 동시 Gate", "size": 11.1, "color": INK, "before": 4},
        {"text": "③ 피크 현금을 시스템·소재 해자로 전환", "size": 11.1, "color": INK, "before": 4},
    ])

    metrics = [
        ("2029", "SK하이닉스 D램·HBM\n생산물량 사실상 판매", CYAN, "[S15]"),
        ("84.4조", "삼성 2026 시설투자\n전년 대비 약 60%↑", BLUE, "[S16]"),
        ("343대", "중국 액침 DUV 확보\n약 270대는 ASML", ORANGE, "[S02]"),
        ("32GW", "2028년 美 DC 전력부족\n필요 전력의 34%", YELLOW, "[S02]"),
    ]
    for i, (big, label, c, src) in enumerate(metrics):
        x = 0.65 + i*3.01
        add_box(slide, x, 3.94, 2.86, 1.60, fill=PANEL, line=GRID)
        add_line(slide, x+0.20, 4.18, x+0.20, 5.27, c, 3)
        add_text(slide, x+0.40, 4.12, 2.15, 0.44, big, size=21.5, color=c, bold=True, font=FONT_EN)
        add_text(slide, x+0.40, 4.62, 2.12, 0.52, label, size=9.8, color=INK, bold=True)
        add_text(slide, x+2.27, 5.16, 0.34, 0.14, src, size=6.8, color=DIM, align=PP_ALIGN.RIGHT)

    add_box(slide, 0.65, 5.82, 12.03, 0.82, fill=PANEL_3, line=None)
    add_pill(slide, 0.88, 6.06, 1.40, 0.30, "CEO TAKEAWAY", fill=CYAN, color=BG, size=9)
    add_text(slide, 2.52, 5.98, 9.78, 0.42, "2027 예산을 ‘칩 CAPA’가 아니라 고객계약·패키징·전력·소재의 동시 확보 기준으로 재배분해야 합니다.", size=12.9, color=WHITE, bold=True, valign=MSO_ANCHOR.MIDDLE)
    footer(slide, 2, "자료: [S02] Oct 06, [S15] Oct 29, [S16] Oct 30 · 기사 내 전망·회사 발언 포함")
    return slide


def signal_map(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 3, "SIGNALS", "16개 이슈는 4단계로 같은 메시지를 반복했다", "258개 헤드라인의 중복 키워드·날짜별 흐름을 통합")

    add_box(slide, 0.65, 1.55, 3.04, 5.10, fill=PANEL, line=GRID)
    add_pill(slide, 0.91, 1.82, 1.24, 0.29, "EVIDENCE BASE", fill=PANEL_3, color=CYAN, size=8.2)
    add_rich_text(slide, 0.91, 2.28, 2.50, 0.47, [
        {"text": "16", "size": 21, "color": CYAN, "bold": True, "font": FONT_EN},
        {"text": " issues  ·  ", "size": 9, "color": MUTED, "bold": True, "font": FONT_EN},
        {"text": "258", "size": 21, "color": LIME, "bold": True, "font": FONT_EN},
    ])
    add_text(slide, 0.91, 2.77, 2.48, 0.25, "OCT 02–30 · HEADLINES", size=8.3, color=DIM, bold=True, font=FONT_EN)
    add_line(slide, 0.91, 3.23, 3.39, 3.23, GRID, 0.8)
    rates = [("AI / COMPUTE", 35, CYAN), ("GEO / CHINA", 26, ORANGE), ("MEMORY", 21, LIME), ("PROCESS / PKG", 19, PURPLE)]
    for i, (lab, val, c) in enumerate(rates):
        y = 3.55 + i*0.61
        add_text(slide, 0.92, y, 1.47, 0.18, lab, size=7.4, color=c, bold=True, font=FONT_EN)
        add_text(slide, 2.78, y-0.01, 0.50, 0.20, f"{val}%", size=9.2, color=c, bold=True, font=FONT_EN, align=PP_ALIGN.RIGHT)
        add_box(slide, 0.92, y+0.25, 2.36, 0.16, fill=GRID, line=None, radius=False)
        add_box(slide, 0.92, y+0.25, 2.36*val/40, 0.16, fill=c, line=None, radius=False)
    add_text(slide, 0.92, 6.12, 2.46, 0.28, "※ 헤드라인 키워드 중복 집계", size=7.6, color=DIM, align=PP_ALIGN.CENTER)

    waves = [
        ("WAVE 1", "OCT 02 · 06 · 07 · 08", "수요 증명", "메모리 가격·AI 냉각·전력난·HBM4가 호황의 실체를 확인", CYAN),
        ("WAVE 2", "OCT 12 · 13 · 14 · 15 · 16", "스택 확장", "맞춤형 칩·CXL·유리기판·4F²·3D D램으로 경쟁 단위가 확대", BLUE),
        ("WAVE 3", "OCT 21 · 22 · 26 · 27", "공급 Lock-in", "LTA·구형 D램 품귀·장비부족·CAPA 지연이 2028–29 공급권을 결정", LIME),
        ("WAVE 4", "OCT 28 · 29 · 30", "현금의 재배치", "실적·투자 확대와 동시에 패키징·소재·현지화 M&A가 가속", ORANGE),
    ]
    for i, (wave, dates, title, desc, c) in enumerate(waves):
        y = 1.55 + i*1.27
        add_box(slide, 3.98, y, 8.70, 1.03, fill=PANEL_2 if i%2 else PANEL, line=GRID)
        add_number_badge(slide, 4.24, y+0.30, f"{i+1}", c, 0.40)
        add_text(slide, 4.84, y+0.16, 1.02, 0.18, wave, size=7.5, color=c, bold=True, font=FONT_EN)
        add_text(slide, 4.84, y+0.43, 1.34, 0.26, title, size=13.3, color=WHITE, bold=True)
        add_pill(slide, 6.34, y+0.16, 2.57, 0.25, dates, fill=PANEL_3, color=MUTED, size=7.3)
        add_text(slide, 6.34, y+0.50, 5.88, 0.27, desc, size=10.1, color=INK, bold=True)
        add_line(slide, 12.39, y+0.18, 12.39, y+0.84, c, 3)
    footer(slide, 3, "자료: 첨부 PDF 16개 이슈의 목차 헤드라인 258건 키워드 코딩 [S01–S16] · 비율 중복")
    return slide


def architecture(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 4, "ARCHITECTURE", "AI 구매 단위가 ‘GPU’에서 ‘AI Factory’로 확장", "성능/W·가동률·조달 안정성을 전체 스택에서 공동 최적화")

    add_box(slide, 0.65, 1.56, 3.16, 4.70, fill=PANEL, line=GRID)
    add_pill(slide, 0.93, 1.84, 1.14, 0.29, "YESTERDAY", fill=PANEL_3, color=MUTED, size=8.3)
    add_text(slide, 0.93, 2.28, 2.58, 0.72, "GPU 확보가\n구매의 중심", size=18.5, color=WHITE, bold=True)
    add_box(slide, 1.43, 3.22, 1.58, 1.14, fill=PANEL_3, line=BLUE, radius=False, width=1.5)
    add_box(slide, 1.89, 3.54, 0.66, 0.50, fill=BLUE, line=None, radius=False)
    add_text(slide, 1.89, 3.64, 0.66, 0.26, "GPU", size=12.2, color=BG, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
    add_multiline(slide, 0.93, 4.74, 2.64, 0.80, [
        {"text": "• 단품 Benchmark", "size": 10.1, "color": MUTED},
        {"text": "• Spot 조달", "size": 10.1, "color": MUTED, "before": 4},
        {"text": "• 학습 중심", "size": 10.1, "color": MUTED, "before": 4},
    ])

    arrow = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(4.00), Inches(3.23), Inches(0.88), Inches(1.02))
    set_fill(arrow, CYAN); set_line(arrow, None)
    add_text(slide, 4.02, 4.42, 0.82, 0.32, "AGENTIC\nINFERENCE", size=7.8, color=CYAN, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)

    add_box(slide, 5.08, 1.56, 7.60, 4.70, fill=PANEL_2, line=CYAN, width=1.1)
    add_pill(slide, 5.38, 1.84, 1.28, 0.29, "TODAY", fill=CYAN, color=BG, size=8.4)
    add_text(slide, 5.38, 2.27, 6.88, 0.35, "계약 단위 = 칩 + 메모리 + 연결 + 패키징 + 유틸리티", size=17.5, color=WHITE, bold=True)
    layers = [
        ("CONTROL", "CPU · AGENT SCHEDULER", BLUE, "AMD / ARM"),
        ("COMPUTE", "GPU + CUSTOM ASIC", CYAN, "MS 1M chips"),
        ("MEMORY", "HBM4 · DRAM · NAND", LIME, "Custom HBM"),
        ("CONNECT", "SILICON PHOTONICS", ORANGE, "4× density"),
        ("INTEGRATE", "2.5D · GLASS · BOND", PURPLE, "Yield / Scale"),
        ("OPERATE", "POWER · COOLING · CAPITAL", YELLOW, "32GW gap"),
    ]
    for i, (tag, layer, c, note) in enumerate(layers):
        y = 2.81 + i*0.48
        w = 6.46 - i*0.16
        x = 5.52 + i*0.08
        add_box(slide, x, y, w, 0.37, fill=PANEL_3, line=c, radius=False, width=0.9)
        add_text(slide, x+0.12, y+0.08, 0.84, 0.16, tag, size=6.5, color=c, bold=True, font=FONT_EN)
        add_text(slide, x+0.98, y+0.055, 3.12, 0.20, layer, size=9.3, color=WHITE, bold=True, font=FONT_EN)
        add_text(slide, x+w-1.38, y+0.07, 1.22, 0.16, note, size=7.3, color=MUTED, bold=True, font=FONT_EN, align=PP_ALIGN.RIGHT)

    add_box(slide, 0.65, 6.45, 12.03, 0.33, fill=CYAN, line=None, radius=False)
    add_text(slide, 0.90, 6.50, 11.52, 0.20, "CEO KPI  |  칩 성능이 아니라 ‘유효 토큰/원 × 가동률 × 공급 커버리지’", size=10.5, color=BG, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 4, "자료: [S04] Oct 08, [S06] Oct 13, [S07] Oct 14, [S08] Oct 15 · 분석적 재구성")
    return slide


def memory_economics(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 5, "MEMORY", "HBM 증설이 범용 메모리까지 조인다: ‘호황’이 아니라 배분의 문제", "웨이퍼 전환 → 레거시 부족 → 가격 상승 → 장기계약 확대의 순환")

    add_box(slide, 0.65, 1.56, 5.18, 4.90, fill=PANEL, line=GRID)
    add_pill(slide, 0.93, 1.84, 1.31, 0.29, "WAFER MIX", fill=PANEL_3, color=CYAN, size=8.4)
    add_text(slide, 0.93, 2.27, 4.55, 0.32, "D램 웨이퍼의 HBM 투입 비중", size=16.3, color=WHITE, bold=True)
    mix = [("2025", 18, MUTED), ("2026", 22, CYAN), ("2027E", 30, LIME)]
    for i, (yr, val, c) in enumerate(mix):
        y = 2.92 + i*0.70
        add_text(slide, 0.94, y, 0.72, 0.20, yr, size=8.6, color=c, bold=True, font=FONT_EN)
        add_box(slide, 1.76, y+0.01, 2.83, 0.29, fill=GRID, line=None, radius=False)
        add_box(slide, 1.76, y+0.01, 2.83*val/35, 0.29, fill=c, line=None, radius=False)
        add_text(slide, 4.68, y-0.01, 0.55, 0.25, f"{val}%", size=10.3, color=c, bold=True, font=FONT_EN, align=PP_ALIGN.RIGHT)
    add_line(slide, 0.93, 5.08, 5.53, 5.08, GRID, 0.8)
    add_text(slide, 0.93, 5.31, 4.58, 0.66, "HBM이 같은 웨이퍼 면적을 더 많이 소모\n→ DDR4·DDR5·낸드 부족이 동반", size=12.0, color=INK, bold=True, align=PP_ALIGN.CENTER)

    cards = [
        (6.10, 1.56, "2029", "D램·HBM 물량\n사실상 판매", CYAN, "SK 발언 [S15]"),
        (9.42, 1.56, "+55%", "2Q 낸드 가격\n전분기 대비", LIME, "[S05]"),
        (6.10, 3.29, "60%", "삼성 HBM 매출 중\nHBM4 비중", BLUE, "3Q [S16]"),
        (9.42, 3.29, "50%", "SK 2027–28 CAPA\nLTA 확보", PURPLE, "[S15]"),
    ]
    for x, y, big, label, c, src in cards:
        add_box(slide, x, y, 3.08, 1.45, fill=PANEL_2, line=c, width=1.0)
        add_text(slide, x+0.25, y+0.24, 1.25, 0.38, big, size=21.5, color=c, bold=True, font=FONT_EN)
        add_text(slide, x+1.35, y+0.23, 1.44, 0.50, label, size=9.7, color=WHITE, bold=True, align=PP_ALIGN.RIGHT)
        add_text(slide, x+0.25, y+1.02, 2.55, 0.18, src, size=7.0, color=DIM, align=PP_ALIGN.RIGHT)

    add_box(slide, 6.10, 5.02, 6.40, 1.44, fill=PANEL_3, line=None)
    add_pill(slide, 6.36, 5.28, 1.08, 0.28, "IMPLICATION", fill=CYAN, color=BG, size=8.4)
    add_text(slide, 7.70, 5.18, 4.46, 0.63, "HBM 수익률만 보지 말고\n레거시 부족의 기회비용까지 가격화", size=13.0, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 5, "자료: [S05] Oct 12, [S13] Oct 27, [S15] Oct 29, [S16] Oct 30 · E=전망")
    return slide


def capa_clock(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 6, "CAPA CLOCK", "투자는 지금, 공급은 2028 이후: 2년의 실행 공백이 가격을 지지", "클린룸·장비·인력·수율 인증이 현금보다 느리다")

    add_line(slide, 1.00, 3.49, 12.15, 3.49, GRID, 5)
    end = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(12.00), Inches(3.30), Inches(0.46), Inches(0.38))
    set_fill(end, CYAN); set_line(end, None)
    years = [
        (1.05, "2026", "84.4조원", "삼성 시설투자\n반도체 77조", BLUE, "[S16]", True),
        (3.40, "2027", "+10%", "ASML 부품가\nHBM 웨이퍼 30%", ORANGE, "[S05·S13]", False),
        (5.75, "2028", "BOTTLENECK", "장비부족 지속\n신규 CAPA 개방", RED, "[S08·S13]", True),
        (8.10, "2029", "LOCKED", "SK 물량 커밋\nTerafab 목표", LIME, "[S11·S15]", False),
        (10.45, "2030+", "3D", "D램 시제품·\n3D 패키징", PURPLE, "[S05·S16]", True),
    ]
    for x, yr, big, desc, c, src, top in years:
        add_circle(slide, x+0.67, 3.27, 0.44, fill=c, line=BG, width=2)
        y = 1.72 if top else 3.96
        add_line(slide, x+0.89, 2.98 if top else 3.71, x+0.89, 3.27 if top else 3.96, c, 1.2)
        add_box(slide, x, y, 1.96, 1.30, fill=PANEL, line=c, width=1.0)
        add_pill(slide, x+0.15, y+0.14, 0.73, 0.25, yr, fill=c, color=BG, size=7.7)
        add_text(slide, x+0.15, y+0.48, 1.60, 0.25, big, size=12.7, color=c, bold=True, font=FONT_EN)
        add_text(slide, x+0.15, y+0.78, 1.60, 0.37, desc, size=8.2, color=INK, bold=True)
        add_text(slide, x+1.16, y+0.13, 0.60, 0.13, src, size=5.7, color=DIM, align=PP_ALIGN.RIGHT)

    add_box(slide, 0.65, 5.76, 12.03, 0.78, fill=PANEL_2, line=None)
    add_pill(slide, 0.90, 5.99, 1.38, 0.30, "INVESTMENT GAP", fill=RED, color=BG, size=8.3)
    add_text(slide, 2.56, 5.91, 9.69, 0.36, "CAPEX 승인 ≠ 공급 확보  |  고객 LTA·장비 Slot·전력 인입·Ramp 인력을 같은 날 승인", size=13.2, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 6, "자료: [S05] Oct 12, [S08] Oct 15, [S11] Oct 22, [S13] Oct 27, [S15–S16]")
    return slide


def chipflation(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 7, "CHIPFLATION", "공급자의 가격결정력이 기기 수요를 훼손하기 시작", "원가 상승이 메모리에서 후공정·장비부품으로 번지며 가치가 상류로 이동")

    add_box(slide, 0.65, 1.56, 5.32, 4.90, fill=PANEL, line=GRID)
    add_pill(slide, 0.93, 1.84, 1.46, 0.29, "COST ESCALATION", fill=PANEL_3, color=ORANGE, size=8.1)
    costs = [
        ("NAND", "+55%", "2Q QoQ", LIME, 0.55),
        ("PACK / TEST", "+10–30%", "긴급 주문 2×", PURPLE, 0.72),
        ("ASML PARTS", "+10%", "2027.01~", BLUE, 0.42),
        ("HBM", "+15–20%", "D램 계약가 전망", CYAN, 0.62),
    ]
    for i, (lab, big, note, c, scale) in enumerate(costs):
        y = 2.43 + i*0.75
        add_text(slide, 0.94, y, 1.22, 0.20, lab, size=7.8, color=c, bold=True, font=FONT_EN)
        add_box(slide, 2.20, y+0.01, 2.72, 0.31, fill=GRID, line=None, radius=False)
        add_box(slide, 2.20, y+0.01, 2.72*scale, 0.31, fill=c, line=None, radius=False)
        add_text(slide, 2.29, y+0.05, 1.14, 0.18, big, size=9.1, color=BG if scale>0.48 else WHITE, bold=True, font=FONT_EN)
        add_text(slide, 4.10, y-0.01, 0.82, 0.22, note, size=7.6, color=MUTED, bold=True, align=PP_ALIGN.RIGHT)
    add_box(slide, 0.93, 5.63, 4.76, 0.50, fill=PANEL_3, line=None)
    add_text(slide, 1.10, 5.73, 4.40, 0.26, "SUPPLIER MARGIN ↑   /   DEVICE BOM ↑", size=9.1, color=WHITE, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)

    arrow = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(6.16), Inches(3.14), Inches(0.91), Inches(1.14))
    set_fill(arrow, ORANGE); set_line(arrow, None)
    add_text(slide, 6.18, 4.48, 0.85, 0.28, "VALUE\nTRANSFER", size=7.4, color=ORANGE, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)

    add_box(slide, 7.28, 1.56, 5.40, 4.90, fill=PANEL_2, line=GRID)
    add_pill(slide, 7.56, 1.84, 1.50, 0.29, "DEMAND DAMAGE", fill=PANEL_3, color=RED, size=8.2)
    declines = [
        ("-21%", "PC 출하량", "메모리값 급등 영향", RED, "[S12]"),
        ("-20%", "Galaxy A 생산", "칩플레이션 대응", ORANGE, "[S12]"),
        ("-10%", "노트북 출하", "2027년 전망", YELLOW, "[S06]"),
    ]
    for i, (big, lab, note, c, src) in enumerate(declines):
        y = 2.48 + i*0.95
        add_circle(slide, 7.61, y, 0.69, fill=c, line=None)
        add_text(slide, 7.66, y+0.19, 0.59, 0.24, big, size=11.0, color=BG, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
        add_text(slide, 8.55, y+0.02, 1.48, 0.22, lab, size=12.1, color=WHITE, bold=True)
        add_text(slide, 8.55, y+0.36, 2.58, 0.20, note, size=8.4, color=MUTED, bold=True)
        add_text(slide, 11.89, y+0.22, 0.37, 0.14, src, size=6.2, color=DIM, align=PP_ALIGN.RIGHT)
    add_box(slide, 7.56, 5.55, 4.84, 0.58, fill=PANEL_3, line=None)
    add_text(slide, 7.76, 5.67, 4.44, 0.30, "가격 극대화보다 고객의 총수요 탄력성 관리", size=11.1, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 7, "자료: [S03] Oct 07, [S05] Oct 12, [S06] Oct 13, [S07] Oct 14, [S12] Oct 26")
    return slide


def bottleneck_stack(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 8, "BOTTLENECKS", "AI 팩토리의 병목은 직렬: 한 층이 막히면 전체 가동률이 멈춘다", "컴퓨트·패키징·네트워크·전력·자본을 하나의 System KPI로 관리")

    items = [
        (0.65, "01", "COMPUTE", "+50%", "가속기당\n메모리 용량", BLUE, "칩/메모리"),
        (3.08, "02", "PACKAGE", "2.5×", "엔비디아 목표\n기판 CAPA", PURPLE, "기판/본딩"),
        (5.51, "03", "NETWORK", "4×", "AI 네트워크\n대역폭 밀도", ORANGE, "광연결"),
        (7.94, "04", "POWER", "32GW", "2028 美\n전력 부족", YELLOW, "냉각/그리드"),
        (10.37, "05", "CAPITAL", "$63B", "AI 칩 구매\n금융 구조", CYAN, "금융/계약"),
    ]
    for x, num, en, big, label, c, owner in items:
        add_box(slide, x, 1.56, 2.18, 4.64, fill=PANEL, line=GRID)
        add_number_badge(slide, x+0.18, 1.78, num, c, 0.40)
        add_text(slide, x+0.72, 1.84, 1.20, 0.18, en, size=8.0, color=c, bold=True, font=FONT_EN)
        add_line(slide, x+1.08, 2.48, x+1.08, 3.04, c, 4)
        add_circle(slide, x+0.83, 2.85, 0.50, fill=c, line=None)
        add_text(slide, x+0.19, 3.56, 1.80, 0.42, big, size=20.0, color=c, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
        add_text(slide, x+0.22, 4.16, 1.74, 0.66, label, size=10.0, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        add_pill(slide, x+0.45, 5.44, 1.28, 0.29, owner, fill=PANEL_3, color=MUTED, size=7.5)
    add_box(slide, 0.65, 6.42, 12.03, 0.34, fill=CYAN, line=None, radius=False)
    add_text(slide, 0.88, 6.47, 11.56, 0.20, "SYSTEM KPI  |  설치 CAPA가 아니라 ‘전력 인가된 양산 CAPA × 패키징 수율 × 계약 커버리지’", size=10.1, color=BG, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 8, "자료: [S02] Oct 06, [S04] Oct 08, [S07] Oct 14, [S11] Oct 22, [S14] Oct 28")
    return slide


def packaging(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 9, "PACKAGING", "첨단패키징은 보조공정이 아니라 AI 성능의 설계공간", "본딩·기판·광연결·검사가 동시에 가치 풀로 부상")

    # central stack
    add_box(slide, 4.70, 1.72, 3.92, 4.72, fill=PANEL_2, line=GRID)
    add_pill(slide, 5.82, 1.96, 1.68, 0.29, "AI PACKAGE STACK", fill=CYAN, color=BG, size=8.2)
    layers = [
        (5.18, 2.52, 2.96, 0.48, "GPU / ASIC", BLUE),
        (5.01, 3.09, 3.30, 0.40, "HBM4 · CUSTOM HBM", LIME),
        (4.88, 3.58, 3.56, 0.35, "HYBRID BOND / TSV", PURPLE),
        (4.82, 4.05, 3.68, 0.49, "INTERPOSER · GLASS / ORGANIC", ORANGE),
        (4.98, 4.69, 3.36, 0.39, "FC-BGA · CCL", CYAN),
    ]
    for x, y, w, h, text, c in layers:
        add_box(slide, x, y, w, h, fill=PANEL_3, line=c, radius=False, width=1.0)
        add_text(slide, x, y+0.09, w, h-0.10, text, size=8.6, color=WHITE, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_line(slide, 5.35, 5.39, 7.98, 5.39, CYAN, 3)
    add_text(slide, 5.10, 5.63, 3.12, 0.38, "Yield × Density × Thermal", size=10.4, color=CYAN, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)

    cards = [
        (0.65, 1.72, "TODAY", "TC / MSVP", "차세대 본더\n2027년 투입", CYAN, "[S08]"),
        (0.65, 4.11, "SCALE", "FC-BGA", "빅테크 LTA·\n베트남 증설", BLUE, "[S02·S07]"),
        (8.91, 1.72, "OPTION", "GLASS / TGV", "2028 양산·\n0PPM 주장", LIME, "[S08·S14]"),
        (8.91, 4.11, "CONTROL", "INSPECTION", "불량 분석\n4h → 30m", ORANGE, "[S07]"),
    ]
    for x, y, tag, title, desc, c, src in cards:
        add_box(slide, x, y, 3.77, 2.02, fill=PANEL, line=c, width=1.0)
        add_pill(slide, x+0.22, y+0.21, 0.88, 0.26, tag, fill=PANEL_3, color=c, size=7.5)
        add_text(slide, x+0.22, y+0.62, 2.82, 0.28, title, size=15.3, color=WHITE, bold=True, font=FONT_EN)
        add_text(slide, x+0.22, y+1.06, 2.55, 0.51, desc, size=10.4, color=INK, bold=True)
        add_text(slide, x+2.86, y+1.58, 0.63, 0.14, src, size=6.1, color=DIM, align=PP_ALIGN.RIGHT)
    footer(slide, 9, "자료: [S02] Oct 06, [S07] Oct 14, [S08] Oct 15, [S14] Oct 28, [S16] Oct 30")
    return slide


def foundry(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 10, "FOUNDRY", "선단 노드만으로는 부족: 고객 생태계·패키징·현지화가 수율을 수익으로 바꾼다")

    data = [
        (0.65, "TSMC", "SCALE LEADER", BLUE, "62.5조", "3Q 매출", ["영업익 36.7조", "A14 SRAM <0.017㎛²", "美 현지 공급망 확대"], "[S09·S13]"),
        (4.72, "SAMSUNG", "TURNAROUND", CYAN, "4nm", "풀가동", ["2nm 빅테크 수주 지연", "자체 AP·HBM 통합", "외부 고객 신뢰 회복"], "[S02·S10·S14]"),
        (8.79, "CHALLENGERS", "ECOSYSTEM BET", ORANGE, "17사", "Rapidus 설계지원", ["日 추가 2,500억엔 검토", "Terafab 2028 Q3 목표", "中 장비·설계 국산화"], "[S02·S10·S11]"),
    ]
    for x, name, tag, c, big, label, bullets, src in data:
        add_box(slide, x, 1.57, 3.89, 4.80, fill=PANEL, line=c, width=1.2)
        add_pill(slide, x+0.24, 1.84, 1.38, 0.29, name, fill=c, color=BG, size=8.5)
        add_text(slide, x+2.01, 1.89, 1.55, 0.18, tag, size=7.2, color=c, bold=True, font=FONT_EN, align=PP_ALIGN.RIGHT)
        add_text(slide, x+0.24, 2.50, 2.15, 0.47, big, size=24, color=c, bold=True, font=FONT_EN)
        add_text(slide, x+2.25, 2.65, 1.23, 0.23, label, size=9.0, color=MUTED, bold=True, align=PP_ALIGN.RIGHT)
        add_line(slide, x+0.24, 3.23, x+3.61, 3.23, GRID, 0.8)
        for i, b in enumerate(bullets):
            yy = 3.55 + i*0.53
            add_circle(slide, x+0.27, yy+0.04, 0.13, fill=c, line=None)
            add_text(slide, x+0.53, yy, 2.86, 0.24, b, size=9.5, color=INK, bold=True)
        add_box(slide, x+0.24, 5.43, 3.37, 0.47, fill=PANEL_3, line=None)
        rule = "현금·수율·패키징" if name=="TSMC" else ("수율·고객·CAPA 균형" if name=="SAMSUNG" else "정부자본·현지수요")
        add_text(slide, x+0.38, 5.52, 3.09, 0.25, rule, size=9.2, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, x+2.82, 6.07, 0.73, 0.13, src, size=5.9, color=DIM, align=PP_ALIGN.RIGHT)
    add_box(slide, 0.65, 6.58, 12.03, 0.25, fill=CYAN, line=None, radius=False)
    add_text(slide, 0.88, 6.61, 11.56, 0.18, "WINNING FORMULA  |  Node PPA × Yield × Customer IP × Advanced Package × Local Access", size=9.7, color=BG, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
    footer(slide, 10)
    return slide


def china_flywheel(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 11, "CHINA", "규제가 만든 역설: 장비 비축과 국산화가 하나의 생태계로 연결")

    add_box(slide, 0.65, 1.58, 8.05, 4.90, fill=PANEL, line=GRID)
    steps = [
        (1.00, 2.17, "01", "EXPORT\nCONTROLS", RED),
        (2.46, 1.82, "02", "DUV\nSTOCKPILE", ORANGE),
        (4.17, 2.14, "03", "LOCAL\nR&D", YELLOW),
        (5.74, 2.96, "04", "DOMESTIC\nSCALE", LIME),
        (4.69, 4.37, "05", "PRICE\nADVANTAGE", CYAN),
        (2.76, 4.54, "06", "MARKET\nSHARE", BLUE),
        (1.19, 3.73, "07", "MORE\nCONTROLS", PURPLE),
    ]
    for i, (x, y, num, label, c) in enumerate(steps):
        add_circle(slide, x, y, 1.12, fill=c, line=None)
        add_text(slide, x+0.06, y+0.18, 1.00, 0.16, num, size=7.2, color=BG, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
        add_text(slide, x+0.10, y+0.43, 0.92, 0.45, label, size=7.2, color=BG, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
        nx, ny, *_ = steps[(i+1)%len(steps)]
        add_line(slide, x+0.56, y+0.56, nx+0.56, ny+0.56, c, 1.3)
    add_circle(slide, 3.47, 3.03, 1.65, fill=PANEL_2, line=CYAN, width=1.4)
    add_text(slide, 3.65, 3.44, 1.29, 0.45, "CHINA\nECOSYSTEM", size=11.2, color=WHITE, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    add_box(slide, 8.98, 1.58, 3.70, 4.90, fill=PANEL_2, line=GRID)
    add_pill(slide, 9.26, 1.84, 1.12, 0.29, "PROOFS", fill=PANEL_3, color=ORANGE, size=8.5)
    proofs = [
        ("343", "액침 DUV 확보", ORANGE, "[S02]"),
        ("34↔28", "NVIDIA vs Huawei", RED, "[S07·S08]"),
        ("#3", "YMTC NAND 출하", LIME, "[S08]"),
        ("4F²", "CXMT 서버 D램", PURPLE, "[S06·S13]"),
        ("18×", "Moore Threads 매출", CYAN, "[S16]"),
    ]
    for i, (big, label, c, src) in enumerate(proofs):
        y = 2.45 + i*0.68
        add_text(slide, 9.27, y, 0.90, 0.29, big, size=15.3, color=c, bold=True, font=FONT_EN)
        add_text(slide, 10.31, y+0.02, 1.67, 0.22, label, size=9.3, color=INK, bold=True)
        add_text(slide, 11.80, y+0.06, 0.58, 0.13, src, size=5.8, color=DIM, align=PP_ALIGN.RIGHT)
    add_box(slide, 9.25, 5.90, 3.14, 0.31, fill=ORANGE, line=None, radius=False)
    add_text(slide, 9.28, 5.56, 3.08, 0.24, "Roadmap ≠ Yield · Claim ≠ Scale", size=8.6, color=MUTED, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
    footer(slide, 11, "자료: [S02] Oct 06, [S03] Oct 07, [S06–S08], [S13], [S16] · 중국 측 계획/주장 포함")
    return slide


def korea_balance(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 12, "KOREA", "K-반도체의 과제: 피크 현금을 ‘메모리 밖의 해자’로 전환")

    add_box(slide, 0.65, 1.57, 5.22, 4.95, fill=PANEL, line=CYAN, width=1.1)
    add_pill(slide, 0.94, 1.85, 1.17, 0.29, "STRENGTH", fill=CYAN, color=BG, size=8.7)
    strengths = [
        ("2029", "D램·HBM 공급권", CYAN, "LTA 기반"),
        (">90%", "양산 수율 역량", LIME, "장비사 평가"),
        ("+222.5%", "10월 1–20일 반도체 수출", BLUE, "YoY"),
        ("+300%", "韓 첨단패키징 CAPA 목표", PURPLE, "2028"),
    ]
    for i, (big, label, c, note) in enumerate(strengths):
        y = 2.40 + i*0.84
        add_text(slide, 0.95, y, 1.24, 0.31, big, size=16.5, color=c, bold=True, font=FONT_EN)
        add_text(slide, 2.30, y+0.01, 2.34, 0.22, label, size=10.2, color=WHITE, bold=True)
        add_pill(slide, 4.75, y, 0.79, 0.24, note, fill=PANEL_3, color=MUTED, size=6.9)
    add_box(slide, 0.94, 5.89, 4.62, 0.34, fill=PANEL_3, line=None)
    add_text(slide, 1.10, 5.95, 4.30, 0.20, "MASS PRODUCTION · MEMORY · PACKAGING", size=8.2, color=CYAN, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)

    add_box(slide, 7.46, 1.57, 5.22, 4.95, fill=PANEL, line=RED, width=1.1)
    add_pill(slide, 7.75, 1.85, 1.17, 0.29, "GAPS", fill=RED, color=BG, size=8.7)
    gaps = [
        ("0%", "AI반도체 예산 집행", RED, "1,666억원"),
        ("IMPORT", "핵심 소재 검증 장벽", ORANGE, "日 의존"),
        ("FOUNDRY", "2nm 고객·생태계", PURPLE, "수주 지연"),
        ("25조", "전기요금 선납 제안", YELLOW, "Grid 부담"),
    ]
    for i, (big, label, c, note) in enumerate(gaps):
        y = 2.40 + i*0.84
        add_text(slide, 7.76, y, 1.24, 0.31, big, size=16.0, color=c, bold=True, font=FONT_EN)
        add_text(slide, 9.10, y+0.01, 2.21, 0.22, label, size=10.2, color=WHITE, bold=True)
        add_pill(slide, 11.45, y, 0.91, 0.24, note, fill=PANEL_3, color=MUTED, size=6.8)
    add_box(slide, 7.75, 5.89, 4.62, 0.34, fill=PANEL_3, line=None)
    add_text(slide, 7.91, 5.95, 4.30, 0.20, "SYSTEM CHIP · MATERIAL · POWER · TALENT", size=8.2, color=RED, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)

    add_circle(slide, 5.82, 2.80, 1.68, fill=PANEL_2, line=CYAN, width=1.5)
    add_text(slide, 6.00, 3.16, 1.32, 0.56, "PEAK\nCASH", size=16.8, color=WHITE, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_line(slide, 5.53, 3.64, 5.82, 3.64, CYAN, 2)
    add_line(slide, 7.50, 3.64, 7.78, 3.64, RED, 2)
    add_pill(slide, 5.66, 4.83, 2.00, 0.29, "CONVERT TO MOAT", fill=PANEL_3, color=LIME, size=8.1)
    footer(slide, 12, "자료: [S01] Oct 02, [S06] Oct 13, [S10–S11], [S15–S16] · 수출 증가율은 기사 인용")
    return slide


def tech_roadmap(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 13, "ROADMAP", "기술 로드맵은 ‘더 작게’와 ‘더 높게’를 동시에 요구", "단면 미세화·3D 적층·광연결·AI 공정제어의 병렬 경쟁")

    years = ["2026", "2027", "2028", "2029", "2030+"]
    for i, yr in enumerate(years):
        x = 2.70 + i*1.95
        add_pill(slide, x, 1.48, 1.45, 0.29, yr, fill=PANEL_3, color=CYAN if i<2 else MUTED, size=8.1)
        if i<4:
            add_line(slide, x+1.45, 1.63, x+1.94, 1.63, GRID, 1.5)

    lanes = [
        ("MEMORY", LIME, ["HBM4 / Custom", "4F² · VCT", "HBM CAPA ramp", "LTA lock", "3D DRAM"]),
        ("FOUNDRY", BLUE, ["2nm ramp", "GAA yield", "A14 design", "1.4nm prep", "SRAM + 3D"]),
        ("PACKAGE", PURPLE, ["2.5D / TSV", "Next bonder", "Glass pilot", "Hybrid bond", "3D stack"]),
        ("CONNECT", ORANGE, ["SiPh standard", "Optical CCL", "4× density", "CPO scale", "Rack fabric"]),
        ("OPERATIONS", CYAN, ["AX defect", "AI process", "Cleanroom -30%", "Digital twin", "Autonomous fab"]),
    ]
    for r, (lane, c, cells) in enumerate(lanes):
        y = 2.05 + r*0.86
        add_pill(slide, 0.65, y+0.20, 1.66, 0.30, lane, fill=PANEL_3, color=c, size=7.7)
        for col, cell in enumerate(cells):
            x = 2.70 + col*1.95
            add_box(slide, x, y, 1.74, 0.68, fill=PANEL if (r+col)%2==0 else PANEL_2, line=GRID)
            add_line(slide, x+0.13, y+0.14, x+0.13, y+0.54, c, 2.5)
            add_text(slide, x+0.28, y+0.18, 1.27, 0.29, cell, size=8.4, color=INK, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_box(slide, 0.65, 6.47, 12.03, 0.32, fill=PANEL_3, line=None, radius=False)
    add_text(slide, 0.88, 6.51, 11.56, 0.20, "GATE  |  Roadmap 공개보다 Yield · Cost/bit · Customer qualification을 분기별 검증", size=10.0, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 13, "자료: [S05–S08], [S11], [S13], [S16] · 일정은 기사 내 기업 로드맵 및 전망")
    return slide


def risk_radar(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 14, "RISK RADAR", "호황의 최대 리스크는 수요 급락보다 ‘동시 제약’과 2028 전환점")

    x0, y0, pw, ph = 0.90, 1.73, 7.30, 4.70
    add_box(slide, x0, y0, pw/2, ph/2, fill="0C2230", line=None, radius=False)
    add_box(slide, x0+pw/2, y0, pw/2, ph/2, fill="2B2224", line=None, radius=False)
    add_box(slide, x0, y0+ph/2, pw/2, ph/2, fill="0A1D2A", line=None, radius=False)
    add_box(slide, x0+pw/2, y0+ph/2, pw/2, ph/2, fill="182534", line=None, radius=False)
    add_line(slide, x0+pw/2, y0, x0+pw/2, y0+ph, GRID, 1)
    add_line(slide, x0, y0+ph/2, x0+pw, y0+ph/2, GRID, 1)
    add_text(slide, 0.66, 3.52, 0.22, 1.05, "영향도", size=9, color=DIM, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_text(slide, 3.39, 6.52, 2.30, 0.22, "발생 가능성  →", size=9, color=DIM, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, x0+pw-1.20, y0+0.14, 0.92, 0.18, "ACT NOW", size=7.8, color=RED, bold=True, font=FONT_EN, align=PP_ALIGN.RIGHT)

    risks = [
        (5.82, 2.27, 0.94, "전력·패키지\n동시 부족", RED),
        (6.52, 3.23, 0.91, "중국\n가격공세", ORANGE),
        (4.49, 2.91, 0.84, "핵심소재\n의존", PURPLE),
        (3.04, 2.31, 0.82, "HBM\n과잉 2029", BLUE),
        (5.44, 4.48, 0.79, "기기수요\n훼손", YELLOW),
        (3.76, 4.71, 0.77, "AI 금융\n레버리지", CYAN),
        (2.02, 4.02, 0.80, "2nm·3D\n수율", LIME),
        (6.70, 4.63, 0.73, "IP·인재\n유출", RED),
    ]
    for x, y, d, label, c in risks:
        add_circle(slide, x, y, d, fill=c, line=None)
        add_text(slide, x+0.05, y+0.10, d-0.10, d-0.17, label, size=7.8, color=BG, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    add_box(slide, 8.55, 1.73, 4.13, 4.70, fill=PANEL, line=GRID)
    add_pill(slide, 8.84, 2.00, 1.54, 0.29, "EARLY WARNINGS", fill=PANEL_3, color=CYAN, size=8.4)
    indicators = [
        ("01", "LTA / CAPA 커버리지", "월", CYAN),
        ("02", "HBM vs Legacy GM", "월", LIME),
        ("03", "패키징 수율·Lead time", "주", PURPLE),
        ("04", "전력 인입 / MW", "월", YELLOW),
        ("05", "중국 DUV 가동·Yield", "분기", ORANGE),
        ("06", "PC·폰 출하 탄력", "월", RED),
        ("07", "AI 고객 Debt / Capex", "분기", BLUE),
    ]
    for i, (num, label, freq, c) in enumerate(indicators):
        yy = 2.45 + i*0.49
        add_number_badge(slide, 8.87, yy, num, c, 0.29)
        add_text(slide, 9.32, yy-0.01, 2.48, 0.19, label, size=9.0, color=INK, bold=True)
        add_pill(slide, 11.88, yy-0.01, 0.50, 0.21, freq, fill=PANEL_3, color=MUTED, size=6.4)
    add_box(slide, 8.84, 5.98, 3.55, 0.24, fill=RED, line=None, radius=False)
    add_text(slide, 8.84, 5.59, 3.55, 0.27, "3개 동시 악화 시 CAPEX Gate 재심의", size=9.2, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 14)
    return slide


def opportunity(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 15, "VALUE POOLS", "가치 풀은 웨이퍼 밖으로: 공급권·통합·운영 최적화가 고마진을 만든다", "버블 크기 = 상대 자본강도 · 위치 = 시장 매력도와 진입 난도(정성 가설)")

    x0, y0, pw, ph = 0.92, 1.74, 8.22, 4.83
    add_box(slide, x0, y0, pw, ph, fill=PANEL, line=GRID, radius=False)
    for frac in [0.33, 0.66]:
        add_line(slide, x0+pw*frac, y0, x0+pw*frac, y0+ph, GRID, 0.65)
        add_line(slide, x0, y0+ph*frac, x0+pw, y0+ph*frac, GRID, 0.65)
    add_text(slide, 0.63, 3.34, 0.22, 1.15, "시장 매력도", size=9, color=DIM, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_text(slide, 3.57, 6.67, 2.90, 0.20, "진입 난도 / CAPEX  →", size=9, color=DIM, bold=True, align=PP_ALIGN.CENTER)
    add_pill(slide, 1.12, 1.94, 1.38, 0.25, "QUICK WINS", fill="12312F", color=CYAN, size=7.8)
    add_pill(slide, 7.39, 1.94, 1.38, 0.25, "BIG BETS", fill="2B2434", color=PURPLE, size=7.8)

    bubbles = [
        (1.88, 2.39, 1.00, "공정 AI·\n검사", CYAN),
        (3.15, 2.08, 1.18, "패키징\n장비", LIME),
        (4.95, 2.42, 1.24, "CUSTOM\nHBM", BLUE),
        (6.86, 2.00, 1.52, "HBM4 /\nNVHBM", PURPLE),
        (2.02, 4.06, 1.02, "냉각·\n전력절감", ORANGE),
        (3.89, 4.05, 1.12, "광연결·\nCCL", CYAN),
        (5.70, 4.35, 1.16, "Glass·\nFC-BGA", LIME),
        (7.43, 4.50, 1.28, "2nm·\n3D DRAM", RED),
        (2.94, 5.37, 0.82, "소재\n국산화", YELLOW),
    ]
    for x, y, d, label, c in bubbles:
        add_circle(slide, x, y, d, fill=c, line=None)
        add_text(slide, x+0.05, y+0.08, d-0.10, d-0.15, label, size=8.1 if d<1.0 else 9.0, color=BG, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    add_box(slide, 9.48, 1.74, 3.20, 4.83, fill=PANEL_2, line=GRID)
    add_pill(slide, 9.76, 2.00, 1.42, 0.29, "ALLOCATION", fill=CYAN, color=BG, size=8.2)
    logic = [
        ("1", "PROTECT", "전력·소재·레거시 공급", ORANGE),
        ("2", "SCALE", "검증된 패키징·HBM", CYAN),
        ("3", "PARTNER", "ASIC·광연결·기판", BLUE),
        ("4", "OPTION", "2nm·Glass·3D D램", LIME),
    ]
    for i, (n, en, ko, c) in enumerate(logic):
        yy = 2.66 + i*0.82
        add_number_badge(slide, 9.78, yy, n, c, 0.38)
        add_text(slide, 10.35, yy-0.01, 1.98, 0.18, en, size=8, color=c, bold=True, font=FONT_EN)
        add_text(slide, 10.35, yy+0.26, 1.98, 0.36, ko, size=9.3, color=INK, bold=True)
    add_line(slide, 9.76, 5.98, 12.38, 5.98, GRID, 0.8)
    add_text(slide, 9.76, 6.07, 2.62, 0.36, "고객·역량·현금흐름을 반영해\n분기마다 재배치", size=8.5, color=MUTED, align=PP_ALIGN.CENTER)
    footer(slide, 15, "자료: [S01–S16] 기반 정성적 전략 평가 · 당사 역량 반영 전 가설")
    return slide


def scenarios(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 16, "SCENARIOS", "2028년이 분기점: 공급 정상화보다 ‘누가 먼저 자격을 얻는가’가 중요")

    data = [
        (0.65, "UPSIDE", "Scarcity to 2030", LIME,
         ["LTA가 2029 이후까지 확장", "전력·패키징 증설 지연", "맞춤형 HBM 가격 프리미엄"],
         "용량예약 · 공동설계 · 가격옵션"),
        (4.72, "BASE", "Tight → Balanced", CYAN,
         ["2028 신규 CAPA 점진 가동", "HBM 성장·레거시 부족 공존", "중국은 저가부터 점유 확대"],
         "고객연계 CAPEX · 믹스 최적화"),
        (8.79, "DOWNSIDE", "Glut + Demand Hit", RED,
         ["중국 공급·범용 가격 공세", "PC·폰 수요 추가 위축", "AI 고객 금융·CAPEX 조정"],
         "Gate 중단 · 현금/재고 · 지역 헤지"),
    ]
    for x, name, title, c, triggers, play in data:
        add_box(slide, x, 1.56, 3.89, 4.72, fill=PANEL, line=c, width=1.2)
        add_pill(slide, x+0.24, 1.83, 1.03, 0.30, name, fill=c, color=BG, size=9)
        add_text(slide, x+0.24, 2.24, 3.34, 0.35, title, size=17.2, color=WHITE, bold=True, font=FONT_EN)
        add_text(slide, x+0.24, 2.79, 0.85, 0.18, "TRIGGERS", size=7.8, color=DIM, bold=True, font=FONT_EN)
        for i, t in enumerate(triggers):
            yy = 3.11 + i*0.52
            add_circle(slide, x+0.27, yy+0.04, 0.13, fill=c, line=None)
            add_text(slide, x+0.54, yy, 2.93, 0.25, t, size=9.8, color=INK, bold=True)
        add_line(slide, x+0.24, 4.77, x+3.60, 4.77, GRID, 0.8)
        add_text(slide, x+0.24, 4.96, 0.90, 0.18, "PLAYBOOK", size=7.8, color=c, bold=True, font=FONT_EN)
        add_box(slide, x+0.24, 5.28, 3.35, 0.55, fill=PANEL_3, line=None)
        add_text(slide, x+0.38, 5.38, 3.07, 0.30, play, size=10.5, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_box(slide, 0.65, 6.49, 12.03, 0.33, fill=CYAN, line=None, radius=False)
    add_text(slide, 0.88, 6.54, 11.56, 0.20, "NO-REGRET  |  LTA · 모듈형 CAPEX · 장비 Slot · 전력 옵션 · Dual source · 다운스트림 탄력성", size=10.0, color=BG, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 16)
    return slide


def decisions(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 17, "CEO DECISIONS", "지금 승인해야 할 6가지 선택", "피크 이익을 2028–30 공급권과 시스템 해자로 전환")

    rows = [
        ("01", "고객", "LTA·공동설계·가격옵션을 묶은 ‘공급권 상품’ 표준화", "LTA Cover", CYAN, "D+30"),
        ("02", "CAPA", "HBM·레거시·파운드리의 기회비용 기반 배분 원칙", "GM / Wafer", BLUE, "D+30"),
        ("03", "병목", "패키징·기판·전력·인력을 웨이퍼와 동시 Gate", "Ready CAPA", LIME, "D+45"),
        ("04", "생태계", "Custom ASIC·광연결·소재의 투자/M&A 후보 확정", "Stack GM", PURPLE, "D+60"),
        ("05", "지정학", "미·중·일 블록별 IP·현지화·Dual source 설계", "Access Score", ORANGE, "D+60"),
        ("06", "자본", "주주환원과 3D·2nm 옵션투자의 현금배분 Guardrail", "FCF / ROIC", YELLOW, "D+90"),
    ]
    for i, (num, area, action, kpi, c, due) in enumerate(rows):
        y = 1.53 + i*0.84
        add_box(slide, 0.65, y, 12.03, 0.70, fill=PANEL, line=GRID)
        add_number_badge(slide, 0.91, y+0.15, num, c, 0.39)
        add_text(slide, 1.49, y+0.14, 0.83, 0.22, area, size=10.8, color=c, bold=True)
        add_text(slide, 2.42, y+0.13, 5.95, 0.26, action, size=11.4, color=WHITE, bold=True)
        add_pill(slide, 8.58, y+0.16, 1.86, 0.27, f"KPI  {kpi}", fill=PANEL_3, color=MUTED, size=7.6)
        add_pill(slide, 10.69, y+0.16, 0.83, 0.27, due, fill=c, color=BG, size=8.1)
        add_pill(slide, 11.73, y+0.16, 0.66, 0.27, "승인", fill=PANEL_3, color=c, size=8.1, line=c)
    add_box(slide, 0.65, 6.68, 12.03, 0.22, fill=PANEL_3, line=None, radius=False)
    add_text(slide, 0.88, 6.70, 11.56, 0.17, "원칙  |  CAPA보다 공급권 · 노드보다 생태계 · 가격보다 총수요 · 피크 이익보다 다음 해자", size=9.2, color=INK, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 17)
    return slide


def plan_90(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 18, "90-DAY PLAN", "결정을 실행으로 바꾸는 90일", "진단 → 공급권 Lock-in → 자원 재배치")

    phase_x = [2.32, 5.70, 9.08]
    phases = [("0–30 DAYS", "DIAGNOSE", CYAN), ("31–60 DAYS", "LOCK-IN", BLUE), ("61–90 DAYS", "EXECUTE", LIME)]
    for i, (days, name, c) in enumerate(phases):
        add_pill(slide, phase_x[i], 1.48, 3.08, 0.31, days, fill=c, color=BG, size=9.2)
        add_text(slide, phase_x[i], 1.86, 3.08, 0.23, name, size=9, color=c, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
        if i<2:
            ar = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(phase_x[i]+3.13), Inches(1.52), Inches(0.31), Inches(0.23))
            set_fill(ar, GRID); set_line(ar, None)

    rows = [("COMMERCIAL", CYAN), ("CAPACITY", BLUE), ("ECOSYSTEM", PURPLE), ("RISK / GEO", ORANGE), ("CAPITAL", YELLOW)]
    content = [
        ["Top 고객 공급권 Gap 진단", "LTA·공동설계 조건 제안", "계약 전환·가격옵션 실행"],
        ["HBM/Legacy/Foundry GM 맵", "장비·전력·수율 4-Gate", "라인 배분·증설 집행"],
        ["ASIC·기판·광연결 후보", "PoC·투자/M&A 실사", "양산 인증·해외 레퍼런스"],
        ["중국·소재·IP 취약점", "Dual source·현지화 합의", "CEO 조기경보 대시보드"],
        ["FCF·ROIC Stress test", "주주환원/옵션 Guardrail", "2027 예산 재배분"],
    ]
    for r, (label, c) in enumerate(rows):
        y = 2.29 + r*0.77
        add_pill(slide, 0.65, y+0.16, 1.35, 0.29, label, fill=PANEL_3, color=c, size=7.4)
        for col in range(3):
            x = phase_x[col]
            add_box(slide, x, y, 3.08, 0.59, fill=PANEL if (r+col)%2==0 else PANEL_2, line=GRID)
            add_circle(slide, x+0.20, y+0.22, 0.15, fill=c, line=None)
            add_text(slide, x+0.50, y+0.14, 2.34, 0.28, content[r][col], size=8.9, color=INK, bold=True, valign=MSO_ANCHOR.MIDDLE)

    add_box(slide, 0.65, 6.35, 12.03, 0.45, fill=PANEL_3, line=None)
    add_text(slide, 0.87, 6.48, 0.72, 0.18, "WEEKLY", size=7.4, color=CYAN, bold=True, font=FONT_EN)
    kpis = ["LTA Cover", "Ready CAPA", "Pkg Yield", "Power MW", "Stack GM", "Access", "FCF"]
    for i, k in enumerate(kpis):
        x = 1.68 + i*1.52
        add_box(slide, x, 6.43, 1.29, 0.24, fill=PANEL, line=GRID)
        add_text(slide, x, 6.46, 1.29, 0.16, k, size=7.0, color=WHITE, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
    footer(slide, 18)
    return slide


def sources(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 19, "APPENDIX", "자료 범위와 해석 원칙", "첨부 PDF 16개 이슈 전체를 직접 추출·분류하고 핵심 수치를 원문 대조")

    rows = [
        ("S01", "Oct 02", "AI 예산 0% · 가격 · 냉각/전력"),
        ("S02", "Oct 06", "전력 32GW · DUV 343 · SiPh"),
        ("S03", "Oct 07", "HBM 가격 · 인재 · 美 관세"),
        ("S04", "Oct 08", "HBM4 · HBF · AI 금융"),
        ("S05", "Oct 12", "VCT 2.6× · NAND +55% · ASML"),
        ("S06", "Oct 13", "MS 1M · 전기 25조 · 4F²"),
        ("S07", "Oct 14", "Pkg 가격 · 中 share · CXL"),
        ("S08", "Oct 15", "Glass · 본더 · ASML shortage"),
        ("S09", "Oct 16", "TSMC 실적 · 中 EUV · HBM4"),
        ("S10", "Oct 21", "3D D램 · e-메탄 · 4nm"),
        ("S11", "Oct 22", "수출 · Terafab · 기판 2.5×"),
        ("S12", "Oct 26", "Custom HBM · Chipflation"),
        ("S13", "Oct 27", "CAPA 2028 · A14 · 4F²/3D"),
        ("S14", "Oct 28", "LTA · $63B · Glass/TGV"),
        ("S15", "Oct 29", "2029 커밋 · Pkg +300%"),
        ("S16", "Oct 30", "84.4조 · HBM4 60% · 소재"),
    ]
    panels = [(0.65, rows[:8]), (6.72, rows[8:])]
    for x, part in panels:
        add_box(slide, x, 1.53, 5.96, 4.84, fill=PANEL, line=GRID)
        add_pill(slide, x+0.27, 1.79, 1.16, 0.29, "SOURCE SET", fill=PANEL_3, color=CYAN, size=8.1)
        for i, (sid, date, topics) in enumerate(part):
            y = 2.28 + i*0.48
            add_pill(slide, x+0.28, y, 0.52, 0.23, sid, fill=PANEL_3, color=CYAN, size=6.9)
            add_text(slide, x+0.97, y-0.01, 0.74, 0.19, date, size=7.9, color=WHITE, bold=True, font=FONT_EN)
            add_text(slide, x+1.78, y-0.01, 3.77, 0.21, topics, size=7.7, color=MUTED, align=PP_ALIGN.RIGHT)
            if i<7:
                add_line(slide, x+0.97, y+0.32, x+5.62, y+0.32, GRID, 0.42)

    add_box(slide, 0.65, 6.57, 12.03, 0.31, fill=CYAN, line=None, radius=False)
    add_text(slide, 0.86, 6.61, 11.60, 0.19, "FACT = 기사 내 발표·수치  |  OUTLOOK = 기업/기관 전망  |  SYNTHESIS = 16개 이슈 공통패턴  |  IMPLICATION = CEO 논의용 가설", size=8.4, color=BG, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 19, "Source set: attached SEMI News for YOU PDFs · Oct 02–30, 2026 · 16 issues / 258 headlines")
    return slide


def build() -> Path:
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    props = prs.core_properties
    props.title = "2026년 10월 세미뉴스 종합 CEO 보고"
    props.subject = "AI 호황의 자원 배분전쟁: 메모리·패키징·전력·자본"
    props.author = "Arena.ai"
    props.keywords = "반도체, HBM4, AI Factory, 첨단패키징, 파운드리, 중국, CEO"
    props.comments = "Built from 16 attached October issues; native editable shapes."

    cover(prs)
    executive(prs)
    signal_map(prs)
    architecture(prs)
    memory_economics(prs)
    capa_clock(prs)
    chipflation(prs)
    bottleneck_stack(prs)
    packaging(prs)
    foundry(prs)
    china_flywheel(prs)
    korea_balance(prs)
    tech_roadmap(prs)
    risk_radar(prs)
    opportunity(prs)
    scenarios(prs)
    decisions(prs)
    plan_90(prs)
    sources(prs)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUT)
    return OUT


if __name__ == "__main__":
    print(build())
