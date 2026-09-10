#!/usr/bin/env python3
"""Build a September 2026 semiconductor-news CEO briefing deck.

The visual language matches the April deck: 16:9, dark executive theme,
native editable PowerPoint shapes, concise insight-led headlines, and an
appendix mapping each dated source file.
"""

from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

from build_ceo_seminews_deck import (
    W, H, BG, BG_ALT, PANEL, PANEL_2, PANEL_3, GRID, WHITE, INK, MUTED, DIM,
    CYAN, BLUE, LIME, ORANGE, YELLOW, RED, PURPLE, FONT_EN,
    rgb, set_fill, set_line, add_box, add_circle, add_line, add_text,
    add_rich_text, add_multiline, add_pill, add_number_badge, add_header,
    add_footer, add_slide_bg,
)

OUT = Path("deliverables/2026_09_세미뉴스_CEO_보고.pptx")
SRC_DEFAULT = "자료: 세미뉴스 2026년 9월호 [S1–S7]"


def footer(slide, number: int, source: str = SRC_DEFAULT):
    add_footer(slide, number, source)


def cover(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)

    # technical grid on the right
    for x in [8.0, 8.55, 9.10, 9.65, 10.20, 10.75, 11.30, 11.85, 12.40]:
        add_line(slide, x, 0.55, x, 6.95, GRID, 0.45)
    for y in [0.70, 1.25, 1.80, 2.35, 2.90, 3.45, 4.00, 4.55, 5.10, 5.65, 6.20, 6.75]:
        add_line(slide, 7.65, y, 12.78, y, GRID, 0.45)

    add_pill(slide, 0.75, 0.62, 2.55, 0.32, "CEO BRIEFING · CONFIDENTIAL", fill=PANEL_2, color=CYAN, size=9.5)
    add_text(slide, 0.75, 1.38, 6.85, 0.58, "2026년 9월 세미뉴스", size=32, color=WHITE, bold=True)
    add_text(slide, 0.75, 2.02, 6.85, 1.22, "AI는 ‘추론 스택’으로 확장\n공급망은 ‘블록’으로 재편", size=29, color=WHITE, bold=True)
    add_line(slide, 0.75, 3.49, 2.30, 3.49, CYAN, 4)
    add_text(
        slide, 0.75, 3.76, 6.25, 0.93,
        "7개 이슈의 신호를 수요·생산·기술·경쟁·지정학으로 통합하고\nCEO가 선점해야 할 가치 풀과 방어해야 할 리스크를 제시합니다.",
        size=14.1, color=MUTED,
    )
    add_text(slide, 0.75, 6.38, 6.20, 0.25, "SEMI NEWS  |  SEP 01 · 02 · 03 · 07 · 08 · 09 · 10", size=10.2, color=CYAN, bold=True, font=FONT_EN)
    add_text(slide, 0.75, 6.69, 5.8, 0.20, "Prepared for CEO · 2026.09", size=9, color=DIM, font=FONT_EN)

    # Editable inference-stack illustration
    cx, cy = 10.25, 3.50
    for d, c, lw in [(4.20, GRID, 1.0), (3.45, PANEL_3, 1.0), (2.65, BLUE, 1.2)]:
        add_circle(slide, cx - d / 2, cy - d / 2, d, fill=None, line=c, width=lw)
    chip = add_box(slide, 9.10, 2.35, 2.30, 2.30, fill=PANEL_2, line=CYAN, radius=True, width=2)
    chip.rotation = 45
    core = add_box(slide, 9.68, 2.93, 1.14, 1.14, fill=CYAN, line=None, radius=False)
    core.rotation = 45
    add_text(slide, 9.40, 3.12, 1.70, 0.55, "INFERENCE", size=12.3, color=BG, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    nodes = [
        (8.12, 1.32, "CPU", BLUE), (9.48, 0.92, "ASIC", CYAN),
        (11.52, 1.15, "HBM", LIME), (12.15, 2.52, "OPTICS", ORANGE),
        (12.02, 4.56, "PACK", YELLOW), (10.95, 5.74, "POWER", PURPLE),
        (8.78, 5.78, "DATA", BLUE), (7.67, 4.36, "SW", CYAN),
    ]
    for nx, ny, label, c in nodes:
        add_circle(slide, nx, ny, 0.35, fill=c, line=None)
        add_text(slide, nx - 0.12, ny + 0.42, 0.59, 0.18, label, size=6.8, color=c, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
        add_line(slide, nx + 0.17, ny + 0.17, cx, cy, c, 0.85)
    add_pill(slide, 8.15, 6.29, 4.55, 0.32, "GPU → FULL STACK → SOVEREIGN SUPPLY", fill=PANEL_2, color=BLUE, size=9.1)
    return slide


def executive(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 2, "EXECUTIVE", "9월의 메시지: AI 수요는 넓어지고, 공급 선택지는 좁아졌다")

    add_box(slide, 0.65, 1.48, 7.65, 2.17, fill=PANEL, line=GRID)
    add_pill(slide, 0.92, 1.74, 1.18, 0.30, "ONE-LINE", fill=PANEL_3, color=LIME, size=9)
    add_rich_text(slide, 0.92, 2.15, 7.05, 0.86, [
        {"text": "추론 중심 전환으로 ", "size": 22.5, "color": WHITE, "bold": True},
        {"text": "CPU·ASIC·HBM·광연결·패키징", "size": 22.5, "color": CYAN, "bold": True},
        {"text": "이 하나의 계약 단위가 되고 있다.", "size": 22.5, "color": WHITE, "bold": True},
    ])
    add_text(slide, 0.92, 3.08, 6.95, 0.31, "동시에 미국·중국 블록화와 내부 CAPA 우선배분이 고객의 선택 가능한 공급망을 축소합니다.", size=11.4, color=MUTED)

    add_box(slide, 8.55, 1.48, 4.13, 2.17, fill=PANEL_2, line=CYAN, width=1.3)
    add_pill(slide, 8.84, 1.74, 1.53, 0.30, "CEO IMPERATIVE", fill=CYAN, color=BG, size=9)
    add_multiline(slide, 8.84, 2.15, 3.46, 1.17, [
        {"text": "‘제품’이 아니라 스택을 선점", "size": 17.5, "color": WHITE, "bold": True},
        {"text": "① 장기계약·공동개발로 수요 고정", "size": 11.3, "color": INK, "before": 8},
        {"text": "② 내부/외부 CAPA 배분 원칙 확립", "size": 11.3, "color": INK, "before": 4},
        {"text": "③ 현지화·소재·인재 리스크 동시 헤지", "size": 11.3, "color": INK, "before": 4},
    ])

    metrics = [
        ("50%+", "삼성 4nm CAPA의\nHBM4 베이스다이 배분", CYAN, "[S4]"),
        ("72.5%", "TSMC 파운드리 점유율\n삼성 5.9% · SMIC 5.4%", BLUE, "[S6]"),
        ("$60B", "퀄컴–AWS 구매 가능액\n10년 장기 협력", LIME, "[S6]"),
        ("99.2%", "중국의 일본산 DCS\n최고 보증금률", ORANGE, "[S5·S6]"),
    ]
    for i, (big, label, c, src) in enumerate(metrics):
        x = 0.65 + i * 3.01
        add_box(slide, x, 3.94, 2.86, 1.60, fill=PANEL, line=GRID)
        add_line(slide, x + 0.20, 4.18, x + 0.20, 5.27, c, 3)
        add_text(slide, x + 0.40, 4.12, 2.15, 0.44, big, size=22, color=c, bold=True, font=FONT_EN)
        add_text(slide, x + 0.40, 4.62, 2.12, 0.52, label, size=9.9, color=INK, bold=True)
        add_text(slide, x + 2.27, 5.16, 0.34, 0.14, src, size=6.8, color=DIM, align=PP_ALIGN.RIGHT)

    add_box(slide, 0.65, 5.82, 12.03, 0.82, fill=PANEL_3, line=None)
    add_pill(slide, 0.88, 6.06, 1.40, 0.30, "CEO TAKEAWAY", fill=CYAN, color=BG, size=9)
    add_text(slide, 2.52, 5.98, 9.78, 0.42, "성장 CAPEX의 기준을 ‘수요량’에서 고객 락인·수율·현지 공급 자격으로 바꿔야 합니다.", size=14.0, color=WHITE, bold=True, valign=MSO_ANCHOR.MIDDLE)
    footer(slide, 2, "자료: [S4] Sep 07, [S5] Sep 08, [S6] Sep 09 · 기사상 계획/전망 포함")
    return slide


def timeline(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 3, "SIGNAL TIMELINE", "7개 이슈가 보여준 전환: 기술 경쟁 → 생산·계약·블록 경쟁", "9월 1–10일의 핵심 뉴스 흐름")

    add_line(slide, 0.95, 3.67, 12.03, 3.67, GRID, 4.5)
    chev = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(11.94), Inches(3.49), Inches(0.49), Inches(0.36))
    set_fill(chev, CYAN); set_line(chev, None)

    items = [
        (0.64, "SEP 01", "설계·낸드", ["AI 설계 자동화", "400단+ / 전력 -25%"], CYAN, True, "[S1]"),
        (2.36, "SEP 02", "패키징·소재", ["5μm 범프", "6N 황산 32만t"], BLUE, False, "[S2]"),
        (4.08, "SEP 03", "증설 경쟁", ["TSMC 수요 1.9×", "Micron HBM 2×"], LIME, True, "[S3]"),
        (5.80, "SEP 07", "HBM4 집중", ["4nm 절반+ 배분", "베이스다이 1.5만장/月"], ORANGE, False, "[S4]"),
        (7.52, "SEP 08", "추론 상용화", ["AWS–Qualcomm", "K-ASIC 반복양산"], PURPLE, True, "[S5]"),
        (9.24, "SEP 09", "계약·점유율", ["최대 $60B / 10년", "TSMC 72.5%"], BLUE, False, "[S6]"),
        (10.96, "SEP 10", "기판·PLP", ["유리기판 핸들링", "대면적 자동화"], CYAN, True, "[S7]"),
    ]
    for x, date, title, bullets, c, top, src in items:
        add_circle(slide, x + 0.57, 3.49, 0.36, fill=c, line=BG, width=2)
        cy = 1.67 if top else 4.02
        if top:
            add_line(slide, x + 0.75, 3.08, x + 0.75, 3.49, c, 1.1)
        else:
            add_line(slide, x + 0.75, 3.85, x + 0.75, 4.02, c, 1.1)
        add_box(slide, x, cy, 1.52, 1.34, fill=PANEL, line=GRID)
        add_pill(slide, x + 0.12, cy + 0.13, 0.72, 0.24, date, fill=c, color=BG, size=7.8)
        add_text(slide, x + 0.12, cy + 0.46, 1.25, 0.24, title, size=10.7, color=WHITE, bold=True)
        add_multiline(slide, x + 0.12, cy + 0.76, 1.25, 0.38, [
            {"text": f"• {b}", "size": 7.2, "color": MUTED, "before": 1} for b in bullets
        ])
        add_text(slide, x + 1.19, cy + 0.13, 0.19, 0.13, src, size=5.9, color=DIM, align=PP_ALIGN.RIGHT)

    add_box(slide, 0.95, 5.94, 11.53, 0.66, fill=PANEL_2, line=None)
    add_text(slide, 1.12, 6.08, 11.19, 0.32, "설계 자동화 → 패키징 미세화 → CAPA 선점 → HBM4 집중 → 추론 스택 장기계약 → 블록형 공급망", size=13.4, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 3)
    return slide


def inference_stack(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 4, "DEMAND SHIFT", "AI 수요의 단위가 GPU에서 ‘추론 시스템’으로 확장", "지속 실행되는 에이전트는 CPU·메모리·스토리지·연결을 동시에 요구")

    # Before
    add_box(slide, 0.65, 1.55, 3.20, 4.62, fill=PANEL, line=GRID)
    add_pill(slide, 0.93, 1.83, 1.20, 0.30, "TRAINING ERA", fill=PANEL_3, color=MUTED, size=8.3)
    add_text(slide, 0.93, 2.29, 2.64, 0.70, "‘GPU 확보’가\n의사결정의 중심", size=18.5, color=WHITE, bold=True)
    # single-chip icon
    add_box(slide, 1.42, 3.25, 1.62, 1.16, fill=PANEL_3, line=BLUE, radius=False, width=1.5)
    add_box(slide, 1.89, 3.57, 0.68, 0.53, fill=BLUE, line=None, radius=False)
    add_text(slide, 1.89, 3.67, 0.68, 0.28, "GPU", size=12.5, color=BG, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
    add_text(slide, 0.93, 4.78, 2.65, 0.73, "• 대규모 일회성 학습\n• 범용 가속기 중심\n• 칩 단품 성능 비교", size=10.4, color=MUTED)

    arrow = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(4.05), Inches(3.20), Inches(0.88), Inches(1.05))
    set_fill(arrow, CYAN); set_line(arrow, None)
    add_text(slide, 4.07, 4.43, 0.82, 0.32, "AGENTIC\nAI", size=8.3, color=CYAN, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)

    # Now stack
    add_box(slide, 5.12, 1.55, 7.56, 4.62, fill=PANEL_2, line=CYAN, width=1.1)
    add_pill(slide, 5.42, 1.83, 1.36, 0.30, "INFERENCE ERA", fill=CYAN, color=BG, size=8.4)
    add_text(slide, 5.42, 2.27, 6.85, 0.35, "성능·비용·전력을 ‘전체 스택’에서 공동 최적화", size=18, color=WHITE, bold=True)

    layers = [
        ("ORCHESTRATE", "CPU · 128 CORE", BLUE, "작업 분배·제어"),
        ("COMPUTE", "CUSTOM ASIC / NPU", CYAN, "서비스별 최적화"),
        ("MEMORY", "HBM4 · DDR · V-NAND", LIME, "대역폭·용량"),
        ("CONNECT", "1.6 Tbps OPTICS", ORANGE, "노드 간 확장"),
        ("INTEGRATE", "2.5D · PLP · TEST", PURPLE, "수율·양산속도"),
    ]
    for i, (tag, layer, c, note) in enumerate(layers):
        y = 2.87 + i * 0.59
        w = 5.85 - i * 0.20
        x = 5.62 + i * 0.10
        add_box(slide, x, y, w, 0.45, fill=PANEL_3, line=c, radius=False, width=1)
        add_text(slide, x + 0.15, y + 0.10, 0.88, 0.17, tag, size=6.8, color=c, bold=True, font=FONT_EN)
        add_text(slide, x + 1.05, y + 0.07, 2.55, 0.23, layer, size=10.1, color=WHITE, bold=True, font=FONT_EN)
        add_text(slide, x + w - 1.55, y + 0.09, 1.35, 0.18, note, size=8.2, color=MUTED, bold=True, align=PP_ALIGN.RIGHT)

    add_box(slide, 0.65, 6.40, 12.03, 0.34, fill=CYAN, line=None, radius=False)
    add_text(slide, 0.88, 6.45, 11.56, 0.20, "CEO 관점  |  경쟁 단위 = 칩 성능이 아니라 ‘토큰당 총비용 × 가동률 × 공급 안정성’", size=10.7, color=BG, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 4, "자료: [S1] Sep 01, [S5] Sep 08, [S6] Sep 09 · 분석적 재구성")
    return slide


def capacity_squeeze(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 5, "CAPACITY", "생산능력 경쟁의 병목: 돈보다 ‘라인·장비·숙련인력’", "증설이 동시에 진행되면서 내부 배분과 양산 실행력이 승부처")

    panels = [
        (0.65, "SAMSUNG 4nm", "내부 배분", CYAN, "30K", "15K+", "월 CAPA", "HBM4 베이스다이"),
        (4.72, "TSMC", "장비·건설", BLUE, "1.0×", "1.9×", "기존 추정", "장비 조달 수요"),
        (8.79, "MICRON HBM", "공격 증설", LIME, "50K", "100K", "전년 수준", "연말 월 CAPA"),
    ]
    for x, name, tag, c, a, b, alabel, blabel in panels:
        add_box(slide, x, 1.57, 3.89, 3.78, fill=PANEL, line=c, width=1.1)
        add_pill(slide, x + 0.24, 1.84, 1.22, 0.29, name, fill=c, color=BG, size=8.4)
        add_text(slide, x + 2.15, 1.88, 1.42, 0.20, tag, size=8.3, color=c, bold=True, align=PP_ALIGN.RIGHT)
        add_rich_text(slide, x + 0.28, 2.43, 3.30, 0.48, [
            {"text": a, "size": 20, "color": MUTED, "bold": True, "font": FONT_EN},
            {"text": "  →  ", "size": 16, "color": DIM, "bold": True},
            {"text": b, "size": 24, "color": c, "bold": True, "font": FONT_EN},
        ], align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.32, 2.96, 1.34, 0.24, alabel, size=8.4, color=DIM, align=PP_ALIGN.CENTER)
        add_text(slide, x + 2.08, 2.96, 1.46, 0.24, blabel, size=8.4, color=INK, bold=True, align=PP_ALIGN.CENTER)
        # gauge
        add_box(slide, x + 0.35, 3.48, 3.18, 0.28, fill=GRID, line=None, radius=False)
        if name == "SAMSUNG 4nm":
            add_box(slide, x + 0.35, 3.48, 1.68, 0.28, fill=c, line=None, radius=False)
            add_text(slide, x + 0.41, 3.53, 1.54, 0.17, "HBM4 50%+", size=7.5, color=BG, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
        else:
            add_box(slide, x + 0.35, 3.48, 3.02, 0.28, fill=c, line=None, radius=False)
        add_line(slide, x + 0.31, 4.10, x + 3.58, 4.10, GRID, 0.8)
        if name == "SAMSUNG 4nm":
            notes = ["외부 고객 CAPA 압박", "수직계열화의 양면성"]
        elif name == "TSMC":
            notes = ["약 20개 팹 동시 건설", "현장·숙련인력 부족"]
        else:
            notes = ["HBM 생산능력 2배", "3사 경쟁 강도 상승"]
        add_multiline(slide, x + 0.35, 4.32, 3.05, 0.61, [
            {"text": f"• {n}", "size": 9.5, "color": MUTED, "before": 2} for n in notes
        ])

    add_box(slide, 0.65, 5.67, 12.03, 0.84, fill=PANEL_2, line=None)
    add_pill(slide, 0.90, 5.91, 1.38, 0.30, "DECISION RULE", fill=CYAN, color=BG, size=8.7)
    add_text(slide, 2.52, 5.83, 9.74, 0.34, "CAPEX 승인 = 고객 커버리지 × 장비 납기 × 숙련인력 × 수율 램프업", size=14.4, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 5, "자료: [S3] Sep 03, [S4] Sep 07 · K=천장/월, 기사 추정치")
    return slide


def market_scoreboard(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 6, "SCOREBOARD", "한국의 메모리 우위와 파운드리 열세가 동시에 선명", "같은 AI 사이클 안에서도 가치 포착력은 사업별로 극단적으로 갈림")

    # Foundry panel
    add_box(slide, 0.65, 1.50, 6.02, 4.90, fill=PANEL, line=GRID)
    add_pill(slide, 0.94, 1.79, 1.22, 0.29, "FOUNDRY · 2Q", fill=BLUE, color=BG, size=8.6)
    add_text(slide, 0.94, 2.25, 5.35, 0.32, "TSMC 독주 · 삼성–SMIC 0.5%p", size=16.8, color=WHITE, bold=True)
    fdata = [("TSMC", 72.5, BLUE), ("SAMSUNG", 5.9, CYAN), ("SMIC", 5.4, ORANGE), ("OTHERS", 16.2, DIM)]
    y = 2.93
    for name, val, c in fdata:
        add_text(slide, 0.95, y, 1.05, 0.19, name, size=8.0, color=c, bold=True, font=FONT_EN)
        add_box(slide, 2.03, y + 0.01, 3.62, 0.25, fill=GRID, line=None, radius=False)
        add_box(slide, 2.03, y + 0.01, max(0.26, 3.62 * val / 75.0), 0.25, fill=c, line=None, radius=False)
        add_text(slide, 5.77, y - 0.01, 0.55, 0.23, f"{val:.1f}%", size=9.5, color=c, bold=True, font=FONT_EN, align=PP_ALIGN.RIGHT)
        y += 0.60
    add_box(slide, 0.94, 5.53, 5.42, 0.53, fill=PANEL_3, line=None)
    add_text(slide, 1.12, 5.64, 5.06, 0.28, "과제  |  2nm 수율 + 고객 생태계 + 외부 CAPA 신뢰", size=10.6, color=WHITE, bold=True, align=PP_ALIGN.CENTER)

    # HBM panel
    add_box(slide, 6.94, 1.50, 5.74, 4.90, fill=PANEL_2, line=GRID)
    add_pill(slide, 7.23, 1.79, 1.22, 0.29, "HBM · 2Q", fill=CYAN, color=BG, size=8.6)
    add_text(slide, 7.23, 2.25, 5.10, 0.32, "K-Memory 82% · 3강 경쟁", size=16.8, color=WHITE, bold=True)
    hdata = [("SK HYNIX", 50, CYAN), ("SAMSUNG", 32, BLUE), ("MICRON", 18, PURPLE)]
    # stacked bar
    start = 7.23
    totalw = 5.06
    for name, val, c in hdata:
        ww = totalw * val / 100
        add_box(slide, start, 2.92, ww, 0.55, fill=c, line=None, radius=False)
        if ww > 0.75:
            add_text(slide, start, 3.06, ww, 0.22, f"{val}%", size=10.5, color=BG, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
        start += ww
    for i, (name, val, c) in enumerate(hdata):
        yy = 3.80 + i * 0.55
        add_circle(slide, 7.28, yy + 0.02, 0.16, fill=c, line=None)
        add_text(slide, 7.59, yy - 0.01, 1.48, 0.20, name, size=8.1, color=c, bold=True, font=FONT_EN)
        add_text(slide, 9.20, yy - 0.04, 0.62, 0.26, f"{val}%", size=12.0, color=WHITE, bold=True, font=FONT_EN)
        if name == "SAMSUNG":
            note = "HBM4 램프업"
        elif name == "MICRON":
            note = "CAPA 2×"
        else:
            note = "선두 방어"
        add_pill(slide, 10.15, yy - 0.04, 1.64, 0.25, note, fill=PANEL_3, color=MUTED, size=7.5)
    add_box(slide, 7.23, 5.53, 5.06, 0.53, fill=PANEL_3, line=None)
    add_text(slide, 7.42, 5.64, 4.70, 0.28, "기회  |  베이스다이·패키징·수율로 격차 재편", size=10.6, color=WHITE, bold=True, align=PP_ALIGN.CENTER)

    footer(slide, 6, "자료: [S3] Sep 03, [S6] Sep 09 · 시장조사 인용 수치")
    return slide


def geopolitics(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 7, "GEO-TECH", "공급망 경쟁은 ‘제품’에서 국가 블록·규제·우회경로로 확장", "시장 접근의 조건이 가격·성능에서 현지화·동맹·컴플라이언스로 이동")

    # block left and right
    add_box(slide, 0.65, 1.56, 4.60, 3.98, fill=PANEL, line=BLUE, width=1.2)
    add_pill(slide, 0.93, 1.84, 1.32, 0.30, "U.S. BLOCK", fill=BLUE, color=BG, size=9)
    add_text(slide, 0.93, 2.27, 3.98, 0.34, "PAX SILICA", size=20, color=WHITE, bold=True, font=FONT_EN)
    add_rich_text(slide, 0.93, 2.72, 3.95, 0.40, [
        {"text": "24개국", "size": 23, "color": BLUE, "bold": True},
        {"text": " + EU", "size": 16, "color": INK, "bold": True},
    ])
    add_multiline(slide, 0.93, 3.30, 3.92, 1.35, [
        {"text": "• 핵심광물·장비·HBM 강점 연결", "size": 10.6, "color": INK},
        {"text": "• 미국 내 생산·인재 생태계 요구", "size": 10.6, "color": INK, "before": 5},
        {"text": "• 동남아 원격접속 등 우회 규제 강화", "size": 10.6, "color": INK, "before": 5},
    ])

    add_box(slide, 8.08, 1.56, 4.60, 3.98, fill=PANEL, line=ORANGE, width=1.2)
    add_pill(slide, 8.36, 1.84, 1.42, 0.30, "CHINA BLOCK", fill=ORANGE, color=BG, size=9)
    add_text(slide, 8.36, 2.27, 3.98, 0.34, "WAICO", size=20, color=WHITE, bold=True, font=FONT_EN)
    add_rich_text(slide, 8.36, 2.72, 3.95, 0.40, [
        {"text": "28개국", "size": 23, "color": ORANGE, "bold": True},
        {"text": " 기술 생태계", "size": 15, "color": INK, "bold": True},
    ])
    add_multiline(slide, 8.36, 3.30, 3.92, 1.35, [
        {"text": "• 제3세계 데이터센터·클라우드 확장", "size": 10.6, "color": INK},
        {"text": "• 국산 DUV·핵심부품 내재화", "size": 10.6, "color": INK, "before": 5},
        {"text": "• 소재 관세로 자국 공급망 보호", "size": 10.6, "color": INK, "before": 5},
    ])

    # middle bridge / fractures
    add_circle(slide, 5.71, 2.47, 1.90, fill=PANEL_2, line=PURPLE, width=1.5)
    add_text(slide, 5.81, 2.74, 1.70, 0.74, "AI\nINFRA", size=18, color=WHITE, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_line(slide, 5.25, 3.42, 5.72, 3.42, BLUE, 2.2)
    add_line(slide, 7.61, 3.42, 8.08, 3.42, ORANGE, 2.2)
    add_pill(slide, 5.35, 4.73, 2.62, 0.30, "LOCALIZE OR LOSE ACCESS", fill=PANEL_3, color=PURPLE, size=8.1)

    signals = [
        ("DUV", "연말 12대 계획", ORANGE, "[S5]"),
        ("DCS", "최고 99.2%", RED, "[S5·S6]"),
        ("SEA", "규제 우회 경로", YELLOW, "[S4]"),
        ("TALENT", "현지 제조인력", CYAN, "[S4]"),
    ]
    for i, (tag, body, c, src) in enumerate(signals):
        x = 0.65 + i * 3.01
        add_box(slide, x, 5.82, 2.86, 0.78, fill=PANEL_2, line=GRID)
        add_text(slide, x + 0.18, 5.96, 0.72, 0.18, tag, size=8.2, color=c, bold=True, font=FONT_EN)
        add_text(slide, x + 0.89, 5.92, 1.55, 0.25, body, size=10.2, color=WHITE, bold=True)
        add_text(slide, x + 2.26, 6.36, 0.38, 0.12, src, size=6.0, color=DIM, align=PP_ALIGN.RIGHT)
    footer(slide, 7, "자료: [S2] Sep 02, [S4] Sep 07, [S5] Sep 08, [S6] Sep 09")
    return slide


def tech_levers(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 8, "TECH LEVERS", "성능 경쟁의 5개 레버: 설계·메모리·CPU·연결·패키징", "모든 레버의 최종 목적은 ‘토큰당 총비용’ 절감")

    levers = [
        (0.65, "01", "DESIGN", "AI 자동화", "설계 SW 작업\n자동화 시험", CYAN, "생산성"),
        (3.08, "02", "MEMORY", "400단+ V-NAND", "전력 소비\n25%+ 절감", LIME, "용량/W"),
        (5.51, "03", "CPU", "ARM N4", "128 core · 성능 2×\n효율 1.25× · BW 1.75×", BLUE, "Orchestration"),
        (7.94, "04", "CONNECT", "OPTICAL I/O", "최대 1.6 Tbps\n공동개발", ORANGE, "Scale-out"),
        (10.37, "05", "PACKAGE", "5μm BUMP", "2028 H2 양산 목표\nHigh-NA 1M wafer", PURPLE, "Density/Yield"),
    ]
    for x, num, en, ttl, desc, c, key in levers:
        add_box(slide, x, 1.56, 2.18, 4.61, fill=PANEL, line=GRID)
        add_number_badge(slide, x + 0.18, 1.78, num, c, 0.40)
        add_text(slide, x + 0.72, 1.84, 1.20, 0.18, en, size=8.2, color=c, bold=True, font=FONT_EN)
        # stylized lever graphic
        add_line(slide, x + 0.55, 2.71, x + 1.63, 2.71, GRID, 5)
        add_circle(slide, x + 0.92, 2.43, 0.56, fill=c, line=None)
        add_text(slide, x + 0.97, 2.56, 0.46, 0.21, num, size=9.5, color=BG, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.17, 3.35, 1.84, 0.44, ttl, size=13.2, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.18, 4.04, 1.82, 0.66, desc, size=9.4, color=MUTED, bold=True, align=PP_ALIGN.CENTER)
        add_pill(slide, x + 0.49, 5.46, 1.20, 0.29, key, fill=PANEL_3, color=c, size=8.0)

    add_box(slide, 0.65, 6.40, 12.03, 0.34, fill=CYAN, line=None, radius=False)
    add_text(slide, 0.88, 6.45, 11.56, 0.20, "SYSTEM KPI  |  성능/W · 메모리 대역폭 · 연결 대역폭 · 패키징 수율 · 개발 Lead Time", size=10.4, color=BG, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 8, "자료: [S1] Sep 01, [S2] Sep 02, [S5] Sep 08, [S6] Sep 09")
    return slide


def k_flywheel(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 9, "K-SEMICONDUCTOR", "9월의 긍정 신호: R&D가 ‘반복 양산·해외 매출’로 전환", "설계–장비–소재–메모리의 상용화 플라이휠")

    # Center flywheel
    add_circle(slide, 4.88, 2.06, 3.55, fill=PANEL_2, line=GRID, width=1.2)
    add_circle(slide, 5.68, 2.86, 1.95, fill=PANEL_3, line=CYAN, width=1.5)
    add_text(slide, 5.85, 3.15, 1.60, 0.75, "K-AI\nSCALE-UP", size=17.5, color=WHITE, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    nodes = [
        (5.47, 2.03, "DESIGN", "500㎟+ AI ASIC", CYAN),
        (7.30, 3.20, "EQUIPMENT", "2.5D 양산 장비", BLUE),
        (5.57, 5.02, "MATERIAL", "6N 황산", ORANGE),
        (3.77, 3.18, "MEMORY", "HBM4 Base Die", LIME),
    ]
    for x, y, tag, note, c in nodes:
        add_circle(slide, x, y, 0.82, fill=c, line=None)
        add_text(slide, x + 0.04, y + 0.19, 0.74, 0.20, tag, size=7.3, color=BG, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
        add_text(slide, x - 0.35, y + 0.90, 1.52, 0.23, note, size=8.0, color=c, bold=True, align=PP_ALIGN.CENTER)
    # curved effect using chevrons
    for (x, y, rot, c) in [(6.60, 2.28, 35, BLUE), (7.22, 4.32, 130, ORANGE), (4.67, 5.06, 220, LIME), (3.77, 2.28, 310, CYAN)]:
        ch = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x), Inches(y), Inches(0.52), Inches(0.34))
        ch.rotation = rot; set_fill(ch, c); set_line(ch, None)

    # Proof cards
    cards = [
        (0.65, "SEMIFIVE", "423억원", "상반기 신규 양산 수주", "해외 비중 45%", CYAN, "[S5]"),
        (9.19, "HANMI", "5종", "2.5D 장비 라인업", "4종 양산 공급", BLUE, "[S2]"),
        (0.65, "KOREA ZINC", "32만t", "반도체 황산 CAPA", "국내 50만t 추진", ORANGE, "[S2]"),
        (9.19, "SAMSUNG", "50%+", "4nm → HBM4", "내재화 스케일", LIME, "[S4]"),
    ]
    ys = [1.75, 1.75, 4.41, 4.41]
    for (card, y) in zip(cards, ys):
        x, tag, big, desc, note, c, src = card
        add_box(slide, x, y, 3.45, 1.91, fill=PANEL, line=GRID)
        add_pill(slide, x + 0.22, y + 0.20, 1.10, 0.26, tag, fill=PANEL_3, color=c, size=7.5)
        add_text(slide, x + 0.22, y + 0.61, 1.75, 0.37, big, size=21, color=c, bold=True, font=FONT_EN)
        add_text(slide, x + 0.22, y + 1.04, 2.75, 0.24, desc, size=10.2, color=WHITE, bold=True)
        add_text(slide, x + 0.22, y + 1.39, 2.75, 0.20, note, size=8.6, color=MUTED, bold=True)
        add_text(slide, x + 3.02, y + 1.55, 0.20, 0.13, src, size=6.0, color=DIM, align=PP_ALIGN.RIGHT)

    add_box(slide, 0.65, 6.57, 12.03, 0.25, fill=CYAN, line=None, radius=False)
    add_text(slide, 0.88, 6.60, 11.56, 0.18, "FLYWHEEL  |  설계 IP → 고객 인증 → 장비·소재 동반 → 반복 양산 → 해외 레퍼런스", size=9.8, color=BG, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 9)
    return slide


def risk_radar(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 10, "RISK RADAR", "호황의 리스크는 ‘수요 급락’보다 실행·배분·블록 충돌", "확률과 충격을 분리하고 7개 선행지표로 관리")

    x0, y0, pw, ph = 0.90, 1.78, 7.35, 4.62
    add_box(slide, x0, y0, pw/2, ph/2, fill="0C2230", line=None, radius=False)
    add_box(slide, x0+pw/2, y0, pw/2, ph/2, fill="2B2224", line=None, radius=False)
    add_box(slide, x0, y0+ph/2, pw/2, ph/2, fill="0A1D2A", line=None, radius=False)
    add_box(slide, x0+pw/2, y0+ph/2, pw/2, ph/2, fill="182534", line=None, radius=False)
    add_line(slide, x0+pw/2, y0, x0+pw/2, y0+ph, GRID, 1)
    add_line(slide, x0, y0+ph/2, x0+pw, y0+ph/2, GRID, 1)
    add_text(slide, 0.66, 3.50, 0.23, 1.10, "영향도", size=9, color=DIM, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    add_text(slide, 3.42, 6.52, 2.30, 0.22, "발생 가능성  →", size=9, color=DIM, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, x0+0.18, y0+0.14, 1.10, 0.18, "HIGH IMPACT", size=7.5, color=MUTED, bold=True, font=FONT_EN)
    add_text(slide, x0+pw-1.46, y0+0.14, 1.20, 0.18, "ACT NOW", size=8, color=RED, bold=True, font=FONT_EN, align=PP_ALIGN.RIGHT)

    risks = [
        (5.72, 2.28, 0.90, "미·중·일\n블록 충돌", RED),
        (6.62, 3.15, 0.93, "외부고객\nCAPA 압박", ORANGE),
        (4.55, 2.95, 0.84, "중국\n추격", PURPLE),
        (3.15, 2.30, 0.84, "CAPEX\n과잉", BLUE),
        (5.60, 4.48, 0.78, "물·전력\n인프라", YELLOW),
        (3.86, 4.68, 0.74, "숙련인력\n부족", CYAN),
        (2.03, 4.00, 0.80, "5μm·2nm\n수율", LIME),
    ]
    for x, y, d, label, c in risks:
        add_circle(slide, x, y, d, fill=c, line=None)
        add_text(slide, x+0.05, y+0.10, d-0.10, d-0.17, label, size=8.2, color=BG, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    add_box(slide, 8.60, 1.78, 4.08, 4.62, fill=PANEL, line=GRID)
    add_pill(slide, 8.90, 2.05, 1.54, 0.29, "EARLY WARNINGS", fill=PANEL_3, color=CYAN, size=8.5)
    indicators = [
        ("01", "삼성–SMIC 점유율 Gap", "분기", BLUE),
        ("02", "HBM4 내부 CAPA 비중", "월간", CYAN),
        ("03", "5μm 범프·2nm 수율", "월간", LIME),
        ("04", "중국 DUV 설치·가동", "분기", ORANGE),
        ("05", "소재 관세·규제 변화", "주간", RED),
        ("06", "물·전력 확보율", "월간", YELLOW),
        ("07", "장기계약 매출 커버", "분기", PURPLE),
    ]
    for i, (num, label, freq, c) in enumerate(indicators):
        yy = 2.50 + i * 0.47
        add_number_badge(slide, 8.92, yy, num, c, 0.29)
        add_text(slide, 9.37, yy-0.01, 2.54, 0.19, label, size=9.0, color=INK, bold=True)
        add_pill(slide, 11.92, yy-0.01, 0.50, 0.21, freq, fill=PANEL_3, color=MUTED, size=6.5)
    add_box(slide, 8.90, 5.95, 3.50, 0.24, fill=RED, line=None, radius=False)
    add_text(slide, 8.90, 5.57, 3.50, 0.27, "2개 이상 동시 악화 시 투자 Gate 재심의", size=9.4, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 10)
    return slide


def opportunity(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 11, "VALUE POOLS", "추론 스택의 가치 풀: 가까운 매출과 장기 옵션을 분리", "버블 크기 = 상대적 자본강도 · 위치 = 시장 매력도와 진입 난도(정성 평가)")

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
        (2.02, 2.43, 0.98, "AI 설계\n자동화", CYAN),
        (3.30, 2.12, 1.18, "패키징\n장비", LIME),
        (5.05, 2.48, 1.20, "CUSTOM\nASIC", BLUE),
        (6.92, 2.02, 1.50, "HBM4\nBase Die", PURPLE),
        (2.10, 4.10, 0.98, "6N 소재·\n물 재이용", ORANGE),
        (4.05, 4.08, 1.12, "광연결\n1.6Tbps", CYAN),
        (5.88, 4.42, 1.14, "유리기판·\nPLP", LIME),
        (7.56, 4.55, 1.28, "2nm·\n5μm", RED),
        (3.02, 5.36, 0.78, "범용\n메모리", YELLOW),
    ]
    for x, y, d, label, c in bubbles:
        add_circle(slide, x, y, d, fill=c, line=None)
        add_text(slide, x+0.05, y+0.08, d-0.10, d-0.15, label, size=8.2 if d<1.0 else 9.1, color=BG, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    add_box(slide, 9.48, 1.74, 3.20, 4.83, fill=PANEL_2, line=GRID)
    add_pill(slide, 9.76, 2.00, 1.42, 0.29, "ALLOCATION LOGIC", fill=CYAN, color=BG, size=8.2)
    logic = [
        ("1", "PROTECT", "소재·물·전력·컴플라이언스", ORANGE),
        ("2", "SCALE", "양산 인증된 장비·ASIC", CYAN),
        ("3", "PARTNER", "HBM·광연결·첨단패키징", BLUE),
        ("4", "OPTION", "유리기판·5μm·High-NA", LIME),
    ]
    for i, (n, en, ko, c) in enumerate(logic):
        yy = 2.66 + i * 0.82
        add_number_badge(slide, 9.78, yy, n, c, 0.38)
        add_text(slide, 10.35, yy-0.01, 1.98, 0.18, en, size=8, color=c, bold=True, font=FONT_EN)
        add_text(slide, 10.35, yy+0.26, 1.98, 0.36, ko, size=9.4, color=INK, bold=True)
    add_line(slide, 9.76, 5.98, 12.38, 5.98, GRID, 0.8)
    add_text(slide, 9.76, 6.03, 2.62, 0.42, "당사 고객·역량·재무를 반영해\n버블 위치를 재평가", size=8.7, color=MUTED, align=PP_ALIGN.CENTER)
    footer(slide, 11, "자료: [S1–S7] 기반 정성적 전략 평가 · 당사 역량 반영 전 가설")
    return slide


def scenarios(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 12, "SCENARIO", "2027년의 차이는 AI 수요보다 ‘계약·CAPA·블록화’가 만든다", "트리거에 따라 가격·투자·현지화 속도를 조정")

    data = [
        (0.65, "UPSIDE", "추론 스택 가속", LIME,
         ["10년형 장기계약 확산", "ASIC·CPU·광연결 동시 성장", "패키징·HBM 타이트 지속"],
         "용량예약 · 공동개발 · 가격 옵션"),
        (4.72, "BASE", "타이트한 다극화", CYAN,
         ["HBM 공급 확대와 수요 균형", "TSMC 독주·K메모리 우위", "현지화 비용 점진 상승"],
         "고객연계 CAPEX · 믹스 개선"),
        (8.79, "DOWNSIDE", "블록 충돌·과잉", RED,
         ["관세·수출통제 상호 확대", "중국 범용 물량 공세", "AI 수익화 지연·CAPEX 조정"],
         "Gate 중단 · 재고/현금 · 지역 헤지"),
    ]
    for x, name, title, c, triggers, play in data:
        add_box(slide, x, 1.56, 3.89, 4.68, fill=PANEL, line=c, width=1.2)
        add_pill(slide, x+0.24, 1.83, 1.03, 0.30, name, fill=c, color=BG, size=9)
        add_text(slide, x+0.24, 2.24, 3.34, 0.35, title, size=18, color=WHITE, bold=True)
        add_text(slide, x+0.24, 2.79, 0.85, 0.18, "TRIGGERS", size=7.8, color=DIM, bold=True, font=FONT_EN)
        for i, t in enumerate(triggers):
            yy = 3.11 + i * 0.52
            add_circle(slide, x+0.27, yy+0.04, 0.13, fill=c, line=None)
            add_text(slide, x+0.54, yy, 2.93, 0.25, t, size=10.0, color=INK, bold=True)
        add_line(slide, x+0.24, 4.77, x+3.60, 4.77, GRID, 0.8)
        add_text(slide, x+0.24, 4.96, 0.90, 0.18, "PLAYBOOK", size=7.8, color=c, bold=True, font=FONT_EN)
        add_box(slide, x+0.24, 5.28, 3.35, 0.55, fill=PANEL_3, line=None)
        add_text(slide, x+0.38, 5.38, 3.07, 0.30, play, size=10.8, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_box(slide, 0.65, 6.48, 12.03, 0.34, fill=CYAN, line=None, radius=False)
    add_text(slide, 0.88, 6.53, 11.56, 0.20, "NO-REGRET  |  장기계약 · 모듈형 CAPEX · 현지화 옵션 · Dual source · 고객/내부 CAPA 원칙", size=10.5, color=BG, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 12)
    return slide


def decisions(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 13, "CEO DECISIONS", "지금 승인해야 할 5가지 선택", "추론 성장·CAPA 배분·블록형 공급망을 한 번에 다루는 의사결정 패키지")

    rows = [
        ("01", "포트폴리오", "GPU 인접이 아닌 ‘추론 스택’ 가치 풀로 자원 이동", "STACK GM", CYAN, "D+30"),
        ("02", "CAPA 배분", "내부 HBM4 vs 외부 파운드리 고객 우선순위 명문화", "고객 유지율", BLUE, "D+30"),
        ("03", "고객계약", "장기구매·공동개발·지분옵션 패키지 표준화", "LTA 커버", LIME, "D+45"),
        ("04", "공급망", "미·중 블록별 현지화·소재·컴플라이언스 옵션", "Access Score", ORANGE, "D+60"),
        ("05", "실행기반", "숙련인력·물·전력·5μm/2nm 수율 Gate 승인", "Ramp KPI", PURPLE, "D+90"),
    ]
    for i, (num, area, action, kpi, c, due) in enumerate(rows):
        y = 1.56 + i * 0.99
        add_box(slide, 0.65, y, 12.03, 0.82, fill=PANEL, line=GRID)
        add_number_badge(slide, 0.91, y+0.21, num, c, 0.40)
        add_text(slide, 1.51, y+0.17, 1.18, 0.22, area, size=11.3, color=c, bold=True)
        add_text(slide, 2.76, y+0.16, 5.50, 0.27, action, size=12.6, color=WHITE, bold=True)
        add_pill(slide, 8.54, y+0.20, 1.92, 0.28, f"KPI  {kpi}", fill=PANEL_3, color=MUTED, size=8)
        add_pill(slide, 10.70, y+0.20, 0.82, 0.28, due, fill=c, color=BG, size=8.4)
        add_pill(slide, 11.73, y+0.20, 0.66, 0.28, "승인", fill=PANEL_3, color=c, size=8.4, line=c)
    add_box(slide, 0.65, 6.62, 12.03, 0.25, fill=PANEL_3, line=None, radius=False)
    add_text(slide, 0.88, 6.64, 11.56, 0.18, "원칙  |  단품보다 스택 · 증설보다 계약 · 내재화보다 고객균형 · 글로벌보다 블록별 실행", size=9.6, color=INK, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 13)
    return slide


def plan_90(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide, alt=True)
    add_header(slide, 14, "90-DAY PLAN", "결정을 실행으로 바꾸는 90일", "진단 → 고객·공급망 Lock-in → 투자·운영 전환")

    phase_x = [2.32, 5.70, 9.08]
    phases = [("0–30 DAYS", "DIAGNOSE", CYAN), ("31–60 DAYS", "LOCK-IN", BLUE), ("61–90 DAYS", "EXECUTE", LIME)]
    for i, (days, name, c) in enumerate(phases):
        add_pill(slide, phase_x[i], 1.48, 3.08, 0.31, days, fill=c, color=BG, size=9.2)
        add_text(slide, phase_x[i], 1.86, 3.08, 0.23, name, size=9, color=c, bold=True, font=FONT_EN, align=PP_ALIGN.CENTER)
        if i < 2:
            ar = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(phase_x[i]+3.13), Inches(1.52), Inches(0.31), Inches(0.23))
            set_fill(ar, GRID); set_line(ar, None)

    rows = [("COMMERCIAL", CYAN), ("CAPACITY", BLUE), ("TECH / PARTNER", PURPLE), ("RISK / ESG", ORANGE)]
    content = [
        ["Top 고객 추론 스택 Gap 진단", "LTA·공동개발 조건 제안", "계약 전환·가격 옵션 실행"],
        ["내부/외부 CAPA 수익성 맵", "장비·인력·수율 3-Gate", "라인 배분·증설 집행"],
        ["ASIC·광연결·PLP 파트너 맵", "PoC·양산 인증 패키지", "반복양산·해외 레퍼런스"],
        ["블록·소재·물/전력 취약점", "현지화·Dual source 합의", "CEO 조기경보 대시보드"],
    ]
    for r, (label, c) in enumerate(rows):
        y = 2.32 + r * 0.90
        add_pill(slide, 0.65, y+0.20, 1.35, 0.30, label, fill=PANEL_3, color=c, size=7.8)
        for col in range(3):
            x = phase_x[col]
            add_box(slide, x, y, 3.08, 0.70, fill=PANEL if (r+col)%2==0 else PANEL_2, line=GRID)
            add_circle(slide, x+0.20, y+0.26, 0.16, fill=c, line=None)
            add_text(slide, x+0.52, y+0.19, 2.30, 0.31, content[r][col], size=9.6, color=INK, bold=True, valign=MSO_ANCHOR.MIDDLE)

    add_box(slide, 0.65, 6.10, 12.03, 0.65, fill=PANEL_3, line=None)
    add_text(slide, 0.88, 6.21, 0.70, 0.18, "WEEKLY KPI", size=7.8, color=CYAN, bold=True, font=FONT_EN)
    kpis = ["STACK 매출", "LTA 커버", "외부고객 유지", "수율/Ramp", "Access Score"]
    for i, k in enumerate(kpis):
        x = 1.78 + i * 2.08
        add_box(slide, x, 6.20, 1.84, 0.30, fill=PANEL, line=GRID)
        add_text(slide, x, 6.24, 1.84, 0.17, k, size=8.4, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 14)
    return slide


def sources(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_bg(slide)
    add_header(slide, 15, "APPENDIX", "자료 범위와 해석 원칙", "수치·일정은 각 세미뉴스에 인용된 기사·기관 전망 기준")

    source_rows = [
        ("S1", "세미뉴스 Sep 01 2026.pdf", "AI 설계 자동화 · 400단 V-NAND · 중국 메모리"),
        ("S2", "세미뉴스 Sep 02 2026.pdf", "2.5D 장비 · 5μm 범프 · 6N 황산 · 미중 블록"),
        ("S3", "세미뉴스 Sep 03 2026.pdf", "TSMC 1.9× · Micron 2× · HBM/DRAM 경쟁"),
        ("S4", "세미뉴스 Sep 07 2026.pdf", "HBM4 베이스다이 · 미국 인재 · 중국 우회"),
        ("S5", "세미뉴스 Sep 08 2026.pdf", "AWS–Qualcomm · K-ASIC · 중국 DUV · Arm N4"),
        ("S6", "세미뉴스 Sep 09 2026.pdf", "파운드리 점유율 · 장기계약 · 물 · High-NA"),
        ("S7", "세미뉴스 Sep 10 2026.pdf", "유리기판 · PLP · 9월 10일 공개 신호"),
    ]
    add_box(slide, 0.65, 1.53, 7.70, 4.94, fill=PANEL, line=GRID)
    add_pill(slide, 0.92, 1.80, 1.12, 0.30, "SOURCE SET", fill=CYAN, color=BG, size=9)
    for i, (sid, name, topics) in enumerate(source_rows):
        y = 2.25 + i * 0.58
        add_pill(slide, 0.94, y, 0.53, 0.25, sid, fill=PANEL_3, color=CYAN, size=7.4)
        add_text(slide, 1.66, y-0.02, 2.93, 0.21, name, size=9.5, color=WHITE, bold=True)
        add_text(slide, 4.55, y-0.01, 3.37, 0.24, topics, size=7.9, color=MUTED, align=PP_ALIGN.RIGHT)
        if i < 6:
            add_line(slide, 1.66, y+0.35, 7.98, y+0.35, GRID, 0.45)

    add_box(slide, 8.68, 1.53, 4.00, 4.94, fill=PANEL_2, line=GRID)
    add_pill(slide, 8.96, 1.80, 1.42, 0.30, "READING GUIDE", fill=PANEL_3, color=LIME, size=8.5)
    principles = [
        ("01", "FACT", "기사 내 수치·발표·계획", CYAN),
        ("02", "OUTLOOK", "기관·업계의 전망", BLUE),
        ("03", "SYNTHESIS", "7개 이슈 간 공통 패턴", PURPLE),
        ("04", "IMPLICATION", "일반화한 CEO 의사결정 가설", ORANGE),
    ]
    for i, (num, en, ko, c) in enumerate(principles):
        y = 2.48 + i * 0.70
        add_number_badge(slide, 8.98, y, num, c, 0.36)
        add_text(slide, 9.52, y-0.01, 0.94, 0.18, en, size=7.8, color=c, bold=True, font=FONT_EN)
        add_text(slide, 10.48, y-0.02, 1.82, 0.25, ko, size=9.4, color=INK, bold=True)
    add_line(slide, 8.96, 5.43, 12.40, 5.43, GRID, 0.8)
    add_text(slide, 8.96, 5.63, 3.44, 0.45, "주의  |  투자액·양산 일정·점유율은\n후속 공시와 최신 데이터로 재검증 필요", size=9.1, color=MUTED, bold=True, align=PP_ALIGN.CENTER)

    add_box(slide, 0.65, 6.66, 12.03, 0.21, fill=CYAN, line=None, radius=False)
    add_text(slide, 0.87, 6.67, 11.58, 0.17, "본 자료의 전략 제안은 회사별 제품·고객·재무 상황을 반영하기 전의 CEO 논의용 가설입니다.", size=8.8, color=BG, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 15, "Source set: requested Sep 2026 files [S1–S7] · matching-date public sources cross-checked")
    return slide


def build() -> Path:
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    props = prs.core_properties
    props.title = "2026년 9월 세미뉴스 종합 CEO 보고"
    props.subject = "추론 스택 확장과 블록형 반도체 공급망"
    props.author = "Arena.ai"
    props.keywords = "반도체, HBM4, AI 추론, ASIC, 파운드리, 공급망, CEO"
    props.comments = "Native-shape, fully editable CEO briefing deck."

    cover(prs)
    executive(prs)
    timeline(prs)
    inference_stack(prs)
    capacity_squeeze(prs)
    market_scoreboard(prs)
    geopolitics(prs)
    tech_levers(prs)
    k_flywheel(prs)
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
