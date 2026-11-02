#!/usr/bin/env python3
"""Create the English animated edition of the October 2026 CEO briefing.

The script rebuilds the Korean source deck, translates every Korean text run
while preserving formatting, and adds restrained PowerPoint-native entrance
animations plus fade slide transitions. Animations reveal logical content
blocks on one click, then sequence automatically.
"""

from pathlib import Path
import re

from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.util import Pt
from PIL import ImageFont

from build_october_ceo_seminews_deck import build as build_ko

OUT = Path("deliverables/October_2026_Semiconductor_News_CEO_Briefing_EN_Animated.pptx")
FONT = "Aptos"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

# Exact run-level translations. Keeping translations at run level preserves
# the emphasis colors and weights of the original native-shape design.
T = {
    "CEO BRIEFING · CONFIDENTIAL": "CEO BRIEFING · ANIMATED",
    "2026년 10월 세미뉴스": "OCTOBER 2026 SEMI NEWS",
    "AI 호황은 ‘칩 부족’을 넘어\n전력·패키징·자본의 배분전쟁으로": "AI BOOM: THE RACE FOR POWER,\nPACKAGING & CAPITAL",
    "16개 이슈·258개 헤드라인을 수요, 메모리, 파운드리,\n첨단패키징, 중국 추격, 인프라 관점으로 통합했습니다.": "An executive synthesis of 16 issues and 258 headlines across demand, memory,\nfoundry, advanced packaging, China's catch-up, and infrastructure.",

    "10월의 결론: 수요 검증은 끝났고, 2028–30 자원 확보전이 시작됐다": "DEMAND IS PROVEN—THE 2028–30 RESOURCE RACE BEGINS",
    "AI 호황의 병목이 ": "The AI bottleneck is shifting from ",
    "GPU → 메모리 → 패키징 → 전력·자본": "GPU → MEMORY → PACKAGING → POWER & CAPITAL",
    "으로 이동한다.": ".",
    "장기계약과 현지 생태계가 가격보다 강한 공급권이 되며, 중국은 규제를 국산화 속도로 전환하고 있습니다.": "Long-term contracts and local ecosystems now secure supply better than price; China is converting controls into localization speed.",
    "‘증설’이 아니라 공급권을 확보": "Secure supply rights—not just capacity",
    "① LTA·공동설계로 수요 선점": "1  Lock demand through LTAs and co-design",
    "② 웨이퍼–패키지–전력 동시 Gate": "2  Gate wafers, packaging and power together",
    "③ 피크 현금을 시스템·소재 해자로 전환": "3  Convert peak cash into system and material moats",
    "SK하이닉스 D램·HBM\n생산물량 사실상 판매": "SK hynix DRAM & HBM\ncapacity effectively sold",
    "84.4조": "KRW 84.4T",
    "삼성 2026 시설투자\n전년 대비 약 60%↑": "Samsung 2026 capex\n~60% YoY increase",
    "343대": "343 units",
    "중국 액침 DUV 확보\n약 270대는 ASML": "China immersion DUV stock\n~270 from ASML",
    "2028년 美 DC 전력부족\n필요 전력의 34%": "2028 U.S. DC power gap\n34% of required power",
    "2027 예산을 ‘칩 CAPA’가 아니라 고객계약·패키징·전력·소재의 동시 확보 기준으로 재배분해야 합니다.": "Reallocate the 2027 budget from chip capacity alone to simultaneous control of contracts, packaging, power and materials.",
    "자료: [S02] Oct 06, [S15] Oct 29, [S16] Oct 30 · 기사 내 전망·회사 발언 포함": "Source: [S02] Oct 06, [S15] Oct 29, [S16] Oct 30 · Includes forecasts and company statements",

    "16개 이슈는 4단계로 같은 메시지를 반복했다": "16 ISSUES REPEATED ONE MESSAGE IN FOUR WAVES",
    "258개 헤드라인의 중복 키워드·날짜별 흐름을 통합": "Integrated view of overlapping keywords and the date-by-date signal flow across 258 headlines",
    "※ 헤드라인 키워드 중복 집계": "Overlapping headline keyword counts",
    "수요 증명": "DEMAND",
    "메모리 가격·AI 냉각·전력난·HBM4가 호황의 실체를 확인": "Memory pricing, AI cooling, power shortages and HBM4 validate the cycle",
    "스택 확장": "STACK",
    "맞춤형 칩·CXL·유리기판·4F²·3D D램으로 경쟁 단위가 확대": "Custom silicon, CXL, glass substrates, 4F² and 3D DRAM expand the competitive unit",
    "공급 Lock-in": "LOCK-IN",
    "LTA·구형 D램 품귀·장비부족·CAPA 지연이 2028–29 공급권을 결정": "LTAs, legacy DRAM scarcity, tool shortages and capex delays decide 2028–29 access",
    "현금의 재배치": "REALLOCATION",
    "실적·투자 확대와 동시에 패키징·소재·현지화 M&A가 가속": "Record earnings and capex coincide with faster packaging, materials and localization M&A",
    "자료: 첨부 PDF 16개 이슈의 목차 헤드라인 258건 키워드 코딩 [S01–S16] · 비율 중복": "Source: Keyword coding of 258 table-of-contents headlines from 16 attached PDFs [S01–S16] · Overlapping shares",

    "AI 구매 단위가 ‘GPU’에서 ‘AI Factory’로 확장": "THE AI BUYING UNIT EXPANDS FROM GPU TO AI FACTORY",
    "성능/W·가동률·조달 안정성을 전체 스택에서 공동 최적화": "Jointly optimize performance/W, utilization and supply resilience across the full stack",
    "GPU 확보가\n구매의 중심": "GPU access was\nthe buying focus",
    "• 단품 Benchmark": "• Component benchmark",
    "• Spot 조달": "• Spot procurement",
    "• 학습 중심": "• Training-centric",
    "계약 단위 = 칩 + 메모리 + 연결 + 패키징 + 유틸리티": "CONTRACT UNIT = THE FULL AI FACTORY STACK",
    "CEO KPI  |  칩 성능이 아니라 ‘유효 토큰/원 × 가동률 × 공급 커버리지’": "CEO KPI  |  Not chip speed: useful tokens/$ × utilization × supply coverage",
    "자료: [S04] Oct 08, [S06] Oct 13, [S07] Oct 14, [S08] Oct 15 · 분석적 재구성": "Source: [S04] Oct 08, [S06] Oct 13, [S07] Oct 14, [S08] Oct 15 · Analytical synthesis",

    "HBM 증설이 범용 메모리까지 조인다: ‘호황’이 아니라 배분의 문제": "HBM EXPANSION TIGHTENS ALL MEMORY",
    "웨이퍼 전환 → 레거시 부족 → 가격 상승 → 장기계약 확대의 순환": "Wafer conversion → legacy scarcity → price increases → more long-term contracts",
    "D램 웨이퍼의 HBM 투입 비중": "HBM share of DRAM wafer input",
    "HBM이 같은 웨이퍼 면적을 더 많이 소모\n→ DDR4·DDR5·낸드 부족이 동반": "HBM consumes more wafer area per bit\n→ DDR4, DDR5 and NAND tighten together",
    "D램·HBM 물량\n사실상 판매": "DRAM & HBM output\neffectively sold",
    "SK 발언 [S15]": "SK statement [S15]",
    "2Q 낸드 가격\n전분기 대비": "2Q NAND price\nquarter over quarter",
    "삼성 HBM 매출 중\nHBM4 비중": "HBM4 share of\nSamsung HBM revenue",
    "SK 2027–28 CAPA\nLTA 확보": "SK 2027–28 capacity\ncovered by LTAs",
    "HBM 수익률만 보지 말고\n레거시 부족의 기회비용까지 가격화": "Price not only HBM margin,\nbut the opportunity cost of legacy scarcity",
    "자료: [S05] Oct 12, [S13] Oct 27, [S15] Oct 29, [S16] Oct 30 · E=전망": "Source: [S05] Oct 12, [S13] Oct 27, [S15] Oct 29, [S16] Oct 30 · E=estimate",

    "투자는 지금, 공급은 2028 이후: 2년의 실행 공백이 가격을 지지": "INVEST NOW; SUPPLY ARRIVES AFTER 2028",
    "클린룸·장비·인력·수율 인증이 현금보다 느리다": "Cleanrooms, tools, talent and yield qualification move slower than cash",
    "84.4조원": "KRW 84.4T",
    "삼성 시설투자\n반도체 77조": "Samsung capex\nKRW 77T in semis",
    "ASML 부품가\nHBM 웨이퍼 30%": "ASML parts +10%\nHBM wafers 30%",
    "장비부족 지속\n신규 CAPA 개방": "Tool shortages persist\nnew capacity opens",
    "SK 물량 커밋\nTerafab 목표": "SK output committed\nTerafab target",
    "D램 시제품·\n3D 패키징": "3D DRAM samples\nand 3D packaging",
    "CAPEX 승인 ≠ 공급 확보  |  고객 LTA·장비 Slot·전력 인입·Ramp 인력을 같은 날 승인": "CAPEX APPROVAL ≠ SUPPLY  |  Approve LTAs, tool slots, power and ramp talent on the same day",
    "자료: [S05] Oct 12, [S08] Oct 15, [S11] Oct 22, [S13] Oct 27, [S15–S16]": "Source: [S05] Oct 12, [S08] Oct 15, [S11] Oct 22, [S13] Oct 27, [S15–S16]",

    "공급자의 가격결정력이 기기 수요를 훼손하기 시작": "CHIPFLATION IS STARTING TO DAMAGE DEVICE DEMAND",
    "원가 상승이 메모리에서 후공정·장비부품으로 번지며 가치가 상류로 이동": "Cost inflation is spreading from memory to back-end and tool parts, shifting value upstream",
    "긴급 주문 2×": "Rush orders 2×",
    "D램 계약가 전망": "DRAM contract forecast",
    "PC 출하량": "PC shipments",
    "메모리값 급등 영향": "Memory price impact",
    "Galaxy A 생산": "Galaxy A output",
    "칩플레이션 대응": "Chipflation response",
    "노트북 출하": "Notebook shipments",
    "2027년 전망": "2027 outlook",
    "가격 극대화보다 고객의 총수요 탄력성 관리": "Manage customers' total demand elasticity—not only peak price",
    "자료: [S03] Oct 07, [S05] Oct 12, [S06] Oct 13, [S07] Oct 14, [S12] Oct 26": "Source: [S03] Oct 07, [S05] Oct 12, [S06] Oct 13, [S07] Oct 14, [S12] Oct 26",

    "AI 팩토리의 병목은 직렬: 한 층이 막히면 전체 가동률이 멈춘다": "AI FACTORY BOTTLENECKS ARE SERIAL",
    "컴퓨트·패키징·네트워크·전력·자본을 하나의 System KPI로 관리": "Manage compute, packaging, networking, power and capital as one system KPI",
    "가속기당\n메모리 용량": "Memory per\naccelerator",
    "칩/메모리": "Chip / memory",
    "엔비디아 목표\n기판 CAPA": "NVIDIA target\nsubstrate capacity",
    "기판/본딩": "Substrate / bond",
    "AI 네트워크\n대역폭 밀도": "AI network\nbandwidth density",
    "광연결": "Optical I/O",
    "2028 美\n전력 부족": "2028 U.S.\npower gap",
    "냉각/그리드": "Cooling / grid",
    "AI 칩 구매\n금융 구조": "AI chip purchase\nfinancing",
    "금융/계약": "Finance / contract",
    "SYSTEM KPI  |  설치 CAPA가 아니라 ‘전력 인가된 양산 CAPA × 패키징 수율 × 계약 커버리지’": "SYSTEM KPI  |  Power-ready production capacity × package yield × contract coverage",
    "자료: [S02] Oct 06, [S04] Oct 08, [S07] Oct 14, [S11] Oct 22, [S14] Oct 28": "Source: [S02] Oct 06, [S04] Oct 08, [S07] Oct 14, [S11] Oct 22, [S14] Oct 28",

    "첨단패키징은 보조공정이 아니라 AI 성능의 설계공간": "ADVANCED PACKAGING IS THE AI PERFORMANCE ENGINE",
    "본딩·기판·광연결·검사가 동시에 가치 풀로 부상": "Bonding, substrates, optical I/O and inspection are emerging as value pools",
    "차세대 본더\n2027년 투입": "Next-gen bonder\ndeployed in 2027",
    "빅테크 LTA·\n베트남 증설": "Big Tech LTAs\nand Vietnam scale-up",
    "2028 양산·\n0PPM 주장": "2028 production;\n0-PPM claim",
    "불량 분석\n4h → 30m": "Defect analysis\n4h → 30m",
    "자료: [S02] Oct 06, [S07] Oct 14, [S08] Oct 15, [S14] Oct 28, [S16] Oct 30": "Source: [S02] Oct 06, [S07] Oct 14, [S08] Oct 15, [S14] Oct 28, [S16] Oct 30",

    "선단 노드만으로는 부족: 고객 생태계·패키징·현지화가 수율을 수익으로 바꾼다": "FOUNDRY WINNERS NEED MORE THAN ADVANCED NODES",
    "62.5조": "KRW 62.5T",
    "3Q 매출": "3Q revenue",
    "영업익 36.7조": "KRW 36.7T op. profit",
    "美 현지 공급망 확대": "Expanding U.S. local supply",
    "현금·수율·패키징": "Cash · yield · packaging",
    "풀가동": "fully utilized",
    "2nm 빅테크 수주 지연": "2nm Big Tech wins delayed",
    "자체 AP·HBM 통합": "In-house AP & HBM integration",
    "외부 고객 신뢰 회복": "Rebuild external client trust",
    "수율·고객·CAPA 균형": "Yield · clients · capacity balance",
    "17사": "17 firms",
    "Rapidus 설계지원": "Rapidus design support",
    "日 추가 2,500억엔 검토": "Japan weighs JPY 250B more",
    "Terafab 2028 Q3 목표": "Terafab target: 3Q28",
    "中 장비·설계 국산화": "China localizes tools & design",
    "정부자본·현지수요": "Public capital · local demand",
    "자료: SEMI News for YOU 2026년 10월 16개 이슈 [S01–S16]": "Source: SEMI News for YOU · 16 October 2026 issues [S01–S16]",

    "규제가 만든 역설: 장비 비축과 국산화가 하나의 생태계로 연결": "EXPORT CONTROLS ACCELERATE CHINA'S LOCAL ECOSYSTEM",
    "액침 DUV 확보": "Immersion DUV stock",
    "YMTC NAND 출하": "YMTC NAND shipments",
    "CXMT 서버 D램": "CXMT server DRAM",
    "Moore Threads 매출": "Moore Threads revenue",
    "자료: [S02] Oct 06, [S03] Oct 07, [S06–S08], [S13], [S16] · 중국 측 계획/주장 포함": "Source: [S02] Oct 06, [S03] Oct 07, [S06–S08], [S13], [S16] · Includes Chinese plans and claims",

    "K-반도체의 과제: 피크 현금을 ‘메모리 밖의 해자’로 전환": "K-SEMICONDUCTOR: BUILD MOATS BEYOND MEMORY",
    "D램·HBM 공급권": "DRAM & HBM supply rights",
    "LTA 기반": "LTA-backed",
    "양산 수율 역량": "Mass-production yield",
    "장비사 평가": "Toolmaker assessment",
    "10월 1–20일 반도체 수출": "Semiconductor exports, Oct 1–20",
    "韓 첨단패키징 CAPA 목표": "Korea advanced packaging target",
    "AI반도체 예산 집행": "AI-chip budget executed",
    "1,666억원": "KRW 166.6B",
    "핵심 소재 검증 장벽": "Critical-material qualification barrier",
    "日 의존": "Japan reliance",
    "2nm 고객·생태계": "2nm clients & ecosystem",
    "수주 지연": "Wins delayed",
    "25조": "KRW 25T",
    "전기요금 선납 제안": "Proposed power prepayment",
    "Grid 부담": "Grid burden",
    "자료: [S01] Oct 02, [S06] Oct 13, [S10–S11], [S15–S16] · 수출 증가율은 기사 인용": "Source: [S01] Oct 02, [S06] Oct 13, [S10–S11], [S15–S16] · Export growth as cited",

    "기술 로드맵은 ‘더 작게’와 ‘더 높게’를 동시에 요구": "THE ROADMAP REQUIRES BOTH ‘SMALLER’ AND ‘TALLER’",
    "단면 미세화·3D 적층·광연결·AI 공정제어의 병렬 경쟁": "Parallel races in planar scaling, 3D stacking, optical I/O and AI process control",
    "GATE  |  Roadmap 공개보다 Yield · Cost/bit · Customer qualification을 분기별 검증": "GATE  |  Validate yield · cost/bit · customer qualification quarterly—not roadmap claims",
    "자료: [S05–S08], [S11], [S13], [S16] · 일정은 기사 내 기업 로드맵 및 전망": "Source: [S05–S08], [S11], [S13], [S16] · Timing reflects company roadmaps and forecasts cited",

    "호황의 최대 리스크는 수요 급락보다 ‘동시 제약’과 2028 전환점": "THE BIGGEST RISK: CONCURRENT CONSTRAINTS IN 2028",
    "영향도": "IMPACT",
    "발생 가능성  →": "LIKELIHOOD  →",
    "전력·패키지\n동시 부족": "Power + package\nshortage",
    "중국\n가격공세": "China\nprice pressure",
    "핵심소재\n의존": "Critical material\ndependence",
    "HBM\n과잉 2029": "HBM glut\nin 2029",
    "기기수요\n훼손": "Device demand\ndestruction",
    "AI 금융\n레버리지": "AI financing\nleverage",
    "2nm·3D\n수율": "2nm / 3D\nyield",
    "IP·인재\n유출": "IP / talent\nleakage",
    "LTA / CAPA 커버리지": "LTA / capacity coverage",
    "월": "Monthly",
    "패키징 수율·Lead time": "Package yield / lead time",
    "주": "Weekly",
    "전력 인입 / MW": "Energized power / MW",
    "중국 DUV 가동·Yield": "China DUV utilization / yield",
    "분기": "Quarterly",
    "PC·폰 출하 탄력": "PC / phone shipment elasticity",
    "AI 고객 Debt / Capex": "AI client debt / capex",
    "3개 동시 악화 시 CAPEX Gate 재심의": "Reopen the capex gate if three indicators deteriorate together",

    "가치 풀은 웨이퍼 밖으로: 공급권·통합·운영 최적화가 고마진을 만든다": "VALUE POOLS MOVE BEYOND WAFERS",
    "버블 크기 = 상대 자본강도 · 위치 = 시장 매력도와 진입 난도(정성 가설)": "Bubble size = relative capital intensity · Position = market appeal vs. entry difficulty (qualitative)",
    "시장 매력도": "MARKET APPEAL",
    "진입 난도 / CAPEX  →": "ENTRY DIFFICULTY / CAPEX  →",
    "공정 AI·\n검사": "Process AI /\ninspection",
    "패키징\n장비": "Packaging\ntools",
    "냉각·\n전력절감": "Cooling /\npower savings",
    "광연결·\nCCL": "Optical I/O /\nCCL",
    "소재\n국산화": "Material\nlocalization",
    "전력·소재·레거시 공급": "Power · materials · legacy supply",
    "검증된 패키징·HBM": "Qualified packaging · HBM",
    "ASIC·광연결·기판": "ASIC · optical I/O · substrates",
    "2nm·Glass·3D D램": "2nm · glass · 3D DRAM",
    "고객·역량·현금흐름을 반영해\n분기마다 재배치": "Reallocate quarterly based on\nclients, capabilities and cash flow",
    "자료: [S01–S16] 기반 정성적 전략 평가 · 당사 역량 반영 전 가설": "Source: Qualitative strategic assessment based on [S01–S16] · Pre-capability hypothesis",

    "2028년이 분기점: 공급 정상화보다 ‘누가 먼저 자격을 얻는가’가 중요": "2028 INFLECTION: QUALIFICATION SPEED WINS",
    "LTA가 2029 이후까지 확장": "LTAs extend beyond 2029",
    "전력·패키징 증설 지연": "Power and packaging expansion slips",
    "맞춤형 HBM 가격 프리미엄": "Custom HBM commands a premium",
    "용량예약 · 공동설계 · 가격옵션": "Capacity reservation · co-design · price options",
    "2028 신규 CAPA 점진 가동": "New capacity ramps gradually in 2028",
    "HBM 성장·레거시 부족 공존": "HBM growth and legacy scarcity coexist",
    "중국은 저가부터 점유 확대": "China gains share from lower tiers",
    "고객연계 CAPEX · 믹스 최적화": "Customer-linked capex · mix optimization",
    "중국 공급·범용 가격 공세": "China supply and commodity price pressure",
    "PC·폰 수요 추가 위축": "Further PC and phone demand erosion",
    "AI 고객 금융·CAPEX 조정": "AI clients cut financing and capex",
    "Gate 중단 · 현금/재고 · 지역 헤지": "Pause gates · protect cash/inventory · regional hedge",
    "NO-REGRET  |  LTA · 모듈형 CAPEX · 장비 Slot · 전력 옵션 · Dual source · 다운스트림 탄력성": "NO-REGRET  |  LTAs · modular capex · tool slots · power options · dual source · downstream elasticity",

    "지금 승인해야 할 6가지 선택": "SIX DECISIONS TO APPROVE NOW",
    "피크 이익을 2028–30 공급권과 시스템 해자로 전환": "Convert peak earnings into 2028–30 supply rights and system moats",
    "고객": "CLIENTS",
    "LTA·공동설계·가격옵션을 묶은 ‘공급권 상품’ 표준화": "Standardize a supply-rights offer combining LTAs, co-design and price options",
    "승인": "APPROVE",
    "HBM·레거시·파운드리의 기회비용 기반 배분 원칙": "Allocate HBM, legacy and foundry capacity by opportunity cost",
    "병목": "BOTTLENECKS",
    "패키징·기판·전력·인력을 웨이퍼와 동시 Gate": "Gate packaging, substrates, power and talent together with wafers",
    "생태계": "ECOSYSTEM",
    "Custom ASIC·광연결·소재의 투자/M&A 후보 확정": "Confirm investment/M&A targets in custom ASIC, optical I/O and materials",
    "지정학": "GEO-TECH",
    "미·중·일 블록별 IP·현지화·Dual source 설계": "Design IP, localization and dual-source plans by U.S./China/Japan bloc",
    "자본": "CAPITAL",
    "주주환원과 3D·2nm 옵션투자의 현금배분 Guardrail": "Set cash-allocation guardrails for returns versus 3D and 2nm options",
    "원칙  |  CAPA보다 공급권 · 노드보다 생태계 · 가격보다 총수요 · 피크 이익보다 다음 해자": "PRINCIPLE  |  Access over capacity · ecosystem over node · demand over price · next moat over peak profit",

    "결정을 실행으로 바꾸는 90일": "90 DAYS TO TURN DECISIONS INTO EXECUTION",
    "진단 → 공급권 Lock-in → 자원 재배치": "Diagnose → lock in supply rights → reallocate resources",
    "Top 고객 공급권 Gap 진단": "Diagnose top-client supply-rights gaps",
    "LTA·공동설계 조건 제안": "Propose LTA and co-design terms",
    "계약 전환·가격옵션 실행": "Convert contracts and execute price options",
    "HBM/Legacy/Foundry GM 맵": "Map HBM / legacy / foundry GM",
    "장비·전력·수율 4-Gate": "Four gates: tools, power, talent, yield",
    "라인 배분·증설 집행": "Execute line allocation and expansion",
    "ASIC·기판·광연결 후보": "Map ASIC, substrate and optical targets",
    "PoC·투자/M&A 실사": "PoCs and investment/M&A diligence",
    "양산 인증·해외 레퍼런스": "Production qualification and global references",
    "중국·소재·IP 취약점": "Audit China, material and IP exposure",
    "Dual source·현지화 합의": "Agree dual-source and localization plans",
    "CEO 조기경보 대시보드": "Launch CEO early-warning dashboard",
    "주주환원/옵션 Guardrail": "Set return / option guardrails",
    "2027 예산 재배분": "Reallocate the 2027 budget",

    "자료 범위와 해석 원칙": "SOURCE SCOPE AND INTERPRETATION",
    "첨부 PDF 16개 이슈 전체를 직접 추출·분류하고 핵심 수치를 원문 대조": "All 16 attached PDFs were extracted and classified; key figures were checked against the source text",
    "AI 예산 0% · 가격 · 냉각/전력": "AI budget 0% · pricing · cooling/power",
    "전력 32GW · DUV 343 · SiPh": "Power 32GW · DUV 343 · silicon photonics",
    "HBM 가격 · 인재 · 美 관세": "HBM pricing · talent · U.S. tariffs",
    "HBM4 · HBF · AI 금융": "HBM4 · HBF · AI financing",
    "MS 1M · 전기 25조 · 4F²": "MS 1M · power KRW 25T · 4F²",
    "Pkg 가격 · 中 share · CXL": "Package pricing · China share · CXL",
    "Glass · 본더 · ASML shortage": "Glass · bonders · ASML shortage",
    "TSMC 실적 · 中 EUV · HBM4": "TSMC results · China EUV · HBM4",
    "3D D램 · e-메탄 · 4nm": "3D DRAM · e-methane · 4nm",
    "수출 · Terafab · 기판 2.5×": "Exports · Terafab · substrates 2.5×",
    "2029 커밋 · Pkg +300%": "2029 commitments · packaging +300%",
    "84.4조 · HBM4 60% · 소재": "KRW 84.4T · HBM4 60% · materials",
    "FACT = 기사 내 발표·수치  |  OUTLOOK = 기업/기관 전망  |  SYNTHESIS = 16개 이슈 공통패턴  |  IMPLICATION = CEO 논의용 가설": "FACT = reported data  |  OUTLOOK = company/institution forecast  |  SYNTHESIS = cross-issue pattern  |  IMPLICATION = CEO discussion hypothesis",
}

# Logical reveal regions in inches: x1, y1, x2, y2, effect.
# Content headers and footers stay visible; the main story reveals in 3–6 steps.
REGIONS = {
    1: [(0.5, 1.1, 7.5, 6.9, "fade"), (7.5, 0.8, 13.0, 6.9, "fade")],
    2: [(0.5, 1.4, 8.4, 3.8, "fade"), (8.4, 1.4, 13.0, 3.8, "fade"),
        (0.5, 3.8, 3.6, 5.7, "wipe"), (3.6, 3.8, 6.6, 5.7, "wipe"),
        (6.6, 3.8, 9.6, 5.7, "wipe"), (9.6, 3.8, 13.0, 5.7, "wipe"),
        (0.5, 5.7, 13.0, 6.9, "fade")],
    3: [(0.5, 1.4, 3.8, 6.8, "fade"), (3.8, 1.4, 13.0, 2.8, "wipe"),
        (3.8, 2.8, 13.0, 4.1, "wipe"), (3.8, 4.1, 13.0, 5.4, "wipe"),
        (3.8, 5.4, 13.0, 6.8, "wipe")],
    4: [(0.5, 1.4, 3.9, 6.4, "fade"), (3.9, 2.7, 5.0, 4.6, "wipe"), (5.0, 1.4, 13.0, 6.4, "wipe")],
    5: [(0.5, 1.4, 5.9, 6.6, "fade"), (5.9, 1.4, 9.3, 3.2, "fade"),
        (9.3, 1.4, 13.0, 3.2, "fade"), (5.9, 3.2, 9.3, 5.0, "fade"),
        (9.3, 3.2, 13.0, 5.0, "fade"), (5.9, 5.0, 13.0, 6.6, "fade")],
    6: [(0.5, 1.4, 3.0, 3.4, "fade"), (2.8, 3.4, 5.3, 5.5, "fade"),
        (5.2, 1.4, 7.8, 3.4, "fade"), (7.5, 3.4, 10.0, 5.5, "fade"),
        (9.9, 1.4, 13.0, 3.4, "fade"), (0.5, 5.5, 13.0, 6.8, "wipe")],
    7: [(0.5, 1.4, 6.0, 6.6, "wipe"), (6.0, 2.7, 7.3, 4.7, "wipe"), (7.2, 1.4, 13.0, 6.6, "fade")],
    8: [(0.5, 1.4, 2.9, 6.3, "fade"), (2.9, 1.4, 5.3, 6.3, "fade"),
        (5.3, 1.4, 7.8, 6.3, "fade"), (7.8, 1.4, 10.2, 6.3, "fade"),
        (10.2, 1.4, 13.0, 6.3, "fade"), (0.5, 6.3, 13.0, 6.9, "wipe")],
    9: [(0.5, 1.4, 4.5, 4.0, "fade"), (0.5, 4.0, 4.5, 6.6, "fade"),
        (4.4, 1.4, 8.8, 6.6, "wipe"), (8.8, 1.4, 13.0, 4.0, "fade"), (8.8, 4.0, 13.0, 6.6, "fade")],
    10: [(0.5, 1.4, 4.6, 6.5, "fade"), (4.6, 1.4, 8.8, 6.5, "fade"), (8.8, 1.4, 13.0, 6.5, "fade")],
    11: [(0.5, 1.4, 8.8, 6.6, "wipe"), (8.8, 1.4, 13.0, 6.6, "fade")],
    12: [(0.5, 1.4, 5.8, 6.7, "fade"), (5.7, 2.3, 7.7, 5.4, "fade"), (7.4, 1.4, 13.0, 6.7, "fade")],
    13: [(0.5, 1.4, 13.0, 2.8, "wipe"), (0.5, 2.8, 13.0, 3.7, "wipe"),
         (0.5, 3.7, 13.0, 4.6, "wipe"), (0.5, 4.6, 13.0, 5.5, "wipe"),
         (0.5, 5.5, 13.0, 6.4, "wipe"), (0.5, 6.4, 13.0, 6.9, "fade")],
    14: [(0.5, 1.4, 8.3, 6.6, "fade"), (8.3, 1.4, 13.0, 6.6, "wipe")],
    15: [(0.5, 1.4, 9.3, 6.7, "fade"), (9.3, 1.4, 13.0, 6.7, "wipe")],
    16: [(0.5, 1.4, 4.6, 6.5, "fade"), (4.6, 1.4, 8.8, 6.5, "fade"),
         (8.8, 1.4, 13.0, 6.5, "fade"), (0.5, 6.4, 13.0, 6.9, "wipe")],
    17: [(0.5, 1.4+i*0.84, 13.0, 2.35+i*0.84, "wipe") for i in range(6)],
    18: [(0.5, 1.4, 5.5, 6.3, "fade"), (5.5, 1.4, 8.9, 6.3, "fade"),
         (8.9, 1.4, 13.0, 6.3, "fade"), (0.5, 6.3, 13.0, 6.9, "wipe")],
    19: [(0.5, 1.4, 6.7, 6.5, "fade"), (6.7, 1.4, 13.0, 6.5, "fade"), (0.5, 6.5, 13.0, 6.9, "wipe")],
}


def _translate(prs: Presentation) -> None:
    missing = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    original = run.text
                    if original in T:
                        run.text = T[original]
                    elif re.search(r"[가-힣]", original):
                        missing.append(original)
                    run.font.name = FONT
    if missing:
        unique = "\n".join(repr(x) for x in dict.fromkeys(missing))
        raise RuntimeError(f"Untranslated Korean runs:\n{unique}")


def _wrapped_lines(text: str, font, width_px: float) -> int:
    count = 0
    for explicit in text.split("\n"):
        if not explicit:
            count += 1
            continue
        line = ""
        lines = 1
        for word in explicit.split(" "):
            candidate = word if not line else f"{line} {word}"
            if line and font.getlength(candidate) > width_px:
                lines += 1
                line = word
            else:
                line = candidate
        count += lines
    return count


def _text_ratio(shape, dpi: int = 120) -> float:
    width_px = max(1, shape.width / 914400 * dpi)
    height_px = max(1, shape.height / 914400 * dpi)
    needed = 0.0
    regular = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    bold_font = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
    for paragraph in shape.text_frame.paragraphs:
        if not paragraph.text:
            continue
        sizes = [r.font.size.pt for r in paragraph.runs if r.font.size]
        pt = max(sizes or [12.0])
        bold = any(bool(r.font.bold) for r in paragraph.runs)
        font = ImageFont.truetype(bold_font if bold else regular, max(5, round(pt * dpi / 72)))
        lines = _wrapped_lines(paragraph.text, font, width_px)
        spacing = paragraph.line_spacing if isinstance(paragraph.line_spacing, float) else 1.0
        needed += lines * pt * dpi / 72 * 1.16 * spacing
        if paragraph.space_before:
            needed += paragraph.space_before.pt * dpi / 72
        if paragraph.space_after:
            needed += paragraph.space_after.pt * dpi / 72
    return needed / height_px


def _fit_english_text(prs: Presentation) -> None:
    """Conservatively shrink translated runs that would wrap beyond their boxes."""
    for slide in prs.slides:
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False) or not shape.text.strip():
                continue
            top_in = shape.top / 914400
            height_in = shape.height / 914400
            minimum = 15.5 if top_in < 1.25 and 0.34 <= height_in <= 0.65 else 5.5
            for _ in range(8):
                ratio = _text_ratio(shape)
                if ratio <= 0.94:
                    break
                scale = max(0.80, min(0.95, 0.90 / ratio))
                changed = False
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.font.size and run.font.size.pt > minimum + 0.05:
                            run.font.size = Pt(max(minimum, run.font.size.pt * scale))
                            changed = True
                if not changed:
                    break


def _shape_ids_in_region(slide, region, used):
    x1, y1, x2, y2, _ = region
    ids = []
    for shape in slide.shapes:
        sid = shape.shape_id
        if sid in used:
            continue
        cx = (shape.left + shape.width/2) / 914400
        cy = (shape.top + shape.height/2) / 914400
        # Keep tiny source footnotes and page furniture static.
        if cy >= 6.95 or cy <= 0.72:
            continue
        if x1 <= cx < x2 and y1 <= cy < y2:
            ids.append(sid)
            used.add(sid)
    return ids


def _effect_children(spids, id_counter, effect, duration):
    chunks = []
    for spid in spids:
        sid_set = next(id_counter)
        sid_fx = next(id_counter)
        filter_name = "wipe(fromLeft)" if effect == "wipe" else "fade"
        chunks.append(f"""
          <p:set>
            <p:cBhvr>
              <p:cTn id="{sid_set}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
              <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
              <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
            </p:cBhvr>
            <p:to><p:strVal val="visible"/></p:to>
          </p:set>
          <p:animEffect transition="in" filter="{filter_name}">
            <p:cBhvr><p:cTn id="{sid_fx}" dur="{duration}" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr>
          </p:animEffect>""")
    return "".join(chunks)


def _counter(start=100):
    n = start
    while True:
        yield n
        n += 1


def _add_animations(prs: Presentation) -> None:
    for slide_num, slide in enumerate(prs.slides, 1):
        # Native PowerPoint fade transition.
        transition = parse_xml(f'<p:transition xmlns:p="{P_NS}" spd="med" advClick="1"><p:fade/></p:transition>')
        slide._element.append(transition)

        used = set()
        groups = []
        for region in REGIONS.get(slide_num, []):
            ids = _shape_ids_in_region(slide, region, used)
            if ids:
                groups.append((ids, region[4]))
        if not groups:
            continue

        ids = _counter(100)
        par_xml = []
        grp_id = 1
        for idx, (spids, effect) in enumerate(groups):
            ctn_id = next(ids)
            node_type = "clickEffect" if idx == 0 else "withEffect"
            delay = 0 if idx == 0 else idx * 320
            duration = 520 if effect == "wipe" else 450
            children = _effect_children(spids, ids, effect, duration)
            preset = 22 if effect == "wipe" else 10
            par_xml.append(f"""
              <p:par><p:cTn id="{ctn_id}" presetID="{preset}" presetClass="entr" presetSubtype="0" fill="hold" grpId="{grp_id}" nodeType="{node_type}">
                <p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>
                <p:childTnLst>{children}</p:childTnLst>
              </p:cTn></p:par>""")
            grp_id += 1

        timing_xml = f"""
        <p:timing xmlns:p="{P_NS}">
          <p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
            <p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">
              <p:childTnLst>{''.join(par_xml)}</p:childTnLst>
            </p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"/></p:prevCondLst>
            <p:nextCondLst><p:cond evt="onNext" delay="0"/></p:nextCondLst></p:seq></p:childTnLst>
          </p:cTn></p:par></p:tnLst>
        </p:timing>"""
        slide._element.append(parse_xml(timing_xml))


def build() -> Path:
    source = build_ko()
    prs = Presentation(str(source))
    _translate(prs)
    _fit_english_text(prs)
    _add_animations(prs)
    props = prs.core_properties
    props.title = "October 2026 Semiconductor News — CEO Briefing (Animated English Edition)"
    props.subject = "The AI resource-allocation race: memory, packaging, power and capital"
    props.comments = "English animated edition; native editable shapes; fade/wipe reveals and fade slide transitions."
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUT)
    return OUT


if __name__ == "__main__":
    print(build())
