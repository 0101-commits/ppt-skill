---
version: "1.0"
name: HCG-Slide-Design-System
type: presentation/slide
canvas: "16:9"
canvasPx: "1280x720"   # 13.333in x 7.5in @ 96dpi. 16:9 전용. (요구사항 7)
unit: px               # 1pt = 1.333px (96dpi 기준). HCG 원본 pt값을 px로 환산해 고정.
fontPolicy: "Pretendard ONLY"   # 본문/제목/숫자/영문 전부 Pretendard. (요구사항 4)
description: >
  ㈜휴먼컨설팅그룹(HCG)의 대기업 HR 컨설팅 제안서 deck 8종을 파싱해 추출한
  하우스 스타일을, 4:3(10.83x7.5in)에서 16:9(1280x720)로 재설계한 슬라이드 디자인 시스템.
  단일 메인 블루(#356CB5) 위에 라이트블루 틴트 면(table fill)·코랄(#F16249) 강조·골드/틸 보조를
  얹는 컨설팅 그레이드 팔레트, 11pt 주력 본문의 고밀도 타이포, 상단 키커+우상단 챕터박스+
  결론형 풀폭 제목+우하단 페이지번호+하단 8pt 출처의 '고정 골격(rigid skeleton)'이 핵심.
  10개 설계 요구사항을 하드 제약으로 인코딩하여 슬라이드 간 위치/폰트/사이즈가 절대 흔들리지 않게 함.

# ─────────────────────────────────────────────────────────────
# COLORS  (출처: 첨부 HCG deck 8종의 theme1.xml + slide XML 직접 파싱)
# 사용 빈도는 8개 deck 전 슬라이드 합산 기준.
# ─────────────────────────────────────────────────────────────
colors:
  # Brand Blue (메인 계열)
  primary: "#356CB5"          # HCG 메인 블루 (858회). 강조 텍스트·핵심 도형·차트 1순위 색
  primary-deep: "#1F4E8C"     # (파생) 다크 밴드/헤더 띠/호버. primary 명도 -20%
  primary-steel: "#336699"    # 딥 스틸블루 (428회). 보조 헤더·2순위 도형
  primary-navy: "#0D1467"     # 네이비 (50회). 표지 모자이크 최암부·강한 대비 텍스트
  primary-soft: "#A1D1F1"     # 라이트 블루 (252회). 칼럼/카드 헤더 면, 차트 2순위
  # Blue Tints (면/표 채움 전용)
  tint-blue-1: "#D4E1F2"      # 표 헤더/짝수행 채움 (52회)
  tint-blue-2: "#B0CEF2"      # 강조 셀/선택 행 (37회)
  tint-blue-3: "#A5C4E1"      # 박스 배경 (49회)
  tint-blue-4: "#EAF2FB"      # (파생) 최연 틴트. 넓은 면 배경
  # Accent / Semantic
  accent-coral: "#F16249"     # 코랄 (514회). Positive 강조·화살표·하이라이트 칩
  accent-gold: "#FFCC66"      # 골드 (theme accent4). 주의/하이라이트 칩
  accent-teal: "#50B8B6"      # 틸 (theme accent6). 4순위 차트·보조 강조
  emphasis-red: "#C00000"     # 강조 레드 (479회). 부정/리스크/감소 수치 (시맨틱 전용)
  positive-green: "#00B050"   # 그린 (24회). 증가/달성 수치 (시맨틱 전용)
  # Neutral / Ink
  ink: "#000000"              # 본문 기본 (44,940회)
  ink-slate: "#263238"        # 슬레이트 잉크 (275회). 차분한 본문/캡션 대안
  text-muted: "#919191"       # 보조 텍스트 (10,492회, lt2). 캡션·출처·비활성
  text-faint: "#B7B7B7"       # (파생) 더 연한 비활성/플레이스홀더
  divider: "#DDDDDD"          # 구분선 (243회). 제목 하단 룰·표 경계
  divider-soft: "#EAEAEA"     # 약한 구분선 (78회)
  surface: "#FFFFFF"          # 캔버스/카드 기본 (16,054회)
  surface-gray: "#F5F6F8"     # (파생) 카드/패널 연그레이 배경
  # Brand mark
  hcg-maroon: "#921F0B"       # HCG 로고 마룬 (435회). 페이지번호 옆 워드마크 전용

# ─────────────────────────────────────────────────────────────
# TYPOGRAPHY  (Pretendard 전용 / 1280x720 캔버스 / px = HCG 원본 pt x 1.333)
# 출처: HCG deck 실측 pt 분포(11pt 주력) + pt->px 96dpi 환산
# 각 role의 size/weight는 '슬라이드 불문 고정값'. (요구사항 2)
# ─────────────────────────────────────────────────────────────
typography:
  cover-title:        { fontFamily: Pretendard, fontSize: 38px, fontWeight: 800, lineHeight: 1.22 }  # 표지 제목 (~28pt)
  cover-subtitle:     { fontFamily: Pretendard, fontSize: 20px, fontWeight: 500, lineHeight: 1.40 }  # 표지 부제 (~15pt)
  cover-date:         { fontFamily: Pretendard, fontSize: 15px, fontWeight: 400, lineHeight: 1.40 }  # 표지 일자
  slide-title:        { fontFamily: Pretendard, fontSize: 22px, fontWeight: 700, lineHeight: 1.28 }  # 결론형 제목 (~16pt). 최대 2줄 (요구사항 6)
  kicker:             { fontFamily: Pretendard, fontSize: 18px, fontWeight: 700, lineHeight: 1.20 }  # 좌상단 토픽 라벨 (~13-14pt)
  chapter-indicator:  { fontFamily: Pretendard, fontSize: 13px, fontWeight: 600, lineHeight: 1.10 }  # 우상단 챕터 박스 (~10pt)
  section-header:     { fontFamily: Pretendard, fontSize: 18px, fontWeight: 700, lineHeight: 1.25 }  # 칼럼/밴드 헤더 (~14pt)
  block-header:       { fontFamily: Pretendard, fontSize: 16px, fontWeight: 700, lineHeight: 1.25 }  # 카드 헤더 (~12-13pt)
  body:               { fontFamily: Pretendard, fontSize: 15px, fontWeight: 400, lineHeight: 1.45 }  # 주력 본문 (11pt)
  body-strong:        { fontFamily: Pretendard, fontSize: 15px, fontWeight: 700, lineHeight: 1.45 }  # 본문 강조
  body-sm:            { fontFamily: Pretendard, fontSize: 13px, fontWeight: 400, lineHeight: 1.40 }  # 표/고밀도 본문 (10pt)
  body-sm-strong:     { fontFamily: Pretendard, fontSize: 13px, fontWeight: 700, lineHeight: 1.40 }
  caption:            { fontFamily: Pretendard, fontSize: 12px, fontWeight: 400, lineHeight: 1.35 }  # 캡션 (9pt)
  source-note:        { fontFamily: Pretendard, fontSize: 11px, fontWeight: 400, lineHeight: 1.30 }  # 하단 출처 (8pt) (요구사항 10)
  page-number:        { fontFamily: Pretendard, fontSize: 13px, fontWeight: 600, lineHeight: 1.10 }  # 우하단 페이지 (10pt)
  chart-axis:         { fontFamily: Pretendard, fontSize: 12px, fontWeight: 500, lineHeight: 1.20 }  # 차트 축/범례
  chart-value:        { fontFamily: Pretendard, fontSize: 13px, fontWeight: 700, lineHeight: 1.10 }  # 차트 데이터 라벨
  kpi-number:         { fontFamily: Pretendard, fontSize: 44px, fontWeight: 800, lineHeight: 1.05 }  # 대형 KPI 수치
  kpi-unit:           { fontFamily: Pretendard, fontSize: 16px, fontWeight: 600, lineHeight: 1.10 }  # KPI 단위/라벨

# ─────────────────────────────────────────────────────────────
# LAYOUT GRID  (16:9 / 1280x720 고정 골격)
# 모든 좌표는 px, 좌상단 원점. 아래 좌표는 '슬라이드 불문 고정'. (요구사항 1, 7, 8)
# ─────────────────────────────────────────────────────────────
layout:
  canvas:            { w: 1280, h: 720 }
  safe-margin:       { left: 56, right: 56, top: 28, bottom: 28 }   # content box x:[56,1224] y:[28,692]
  content-width:     1168                                          # 1280 - 56*2
  # 상단 헤더 영역 (헤더 존: y 28~132)
  kicker:            { x: 56,  y: 28,  w: 760, h: 28, align: left,  role: kicker }
  chapter-indicator: { rightEdge: 1224, y: 28, w: 150, h: 26, align: center, role: chapter-indicator, box: true }
  slide-title:       { x: 56,  y: 64,  w: 1168, h: 64, align: left, role: slide-title, maxLines: 2 }  # 결론형 제목
  title-rule:        { x: 56,  y: 136, w: 1168, thickness: 1, color: "{colors.divider}" }            # 제목 하단 1px 룰
  # 본문 영역 (콘텐츠 존: y 152~664) — 차트/도형으로 가득 채움 (요구사항 3,5,9)
  content-area:      { x: 56,  y: 152, w: 1168, h: 512 }
  # 하단 푸터 영역 (푸터 존: y 678~692)
  source-note:       { x: 56,  y: 678, w: 980, h: 14, align: left,  role: source-note }   # 좌하단 출처 (요구사항 10)
  page-number:       { rightEdge: 1224, y: 678, h: 14, align: right, role: page-number }  # 우하단 페이지번호 + HCG마크
  # 표지(cover) 전용 그리드
  cover-hero-band:   { x: 0, y: 0, w: 1280, h: 300 }    # 상단 아이콘 모자이크 밴드
  cover-title:       { x: 160, y: 392, w: 960, h: 56, align: center, role: cover-title }
  cover-subtitle:    { x: 160, y: 470, w: 960, h: 30, align: center, role: cover-subtitle }
  cover-date:        { x: 160, y: 540, w: 960, h: 24, align: center, role: cover-date }
  cover-logo:        { x: 560, y: 612, w: 160, h: 48, align: center }     # HCG 워드마크
  cover-legal:       { x: 160, y: 676, w: 960, h: 28, align: center, role: source-note }  # 저작권 고지

# ─────────────────────────────────────────────────────────────
# SPACING / ROUNDED  (고밀도 유지: 간격은 작고 일관되게. 요구사항 5,9)
# ─────────────────────────────────────────────────────────────
spacing:
  xxs: 4
  xs: 8
  sm: 12
  md: 16
  lg: 20
  xl: 24
  gutter: 24        # 칼럼 사이 표준 간격
  block-gap: 12     # 카드/블록 사이 표준 간격 (0.3"급, 고밀도)
rounded:
  none: 0
  sm: 3
  md: 6
  lg: 10
  card: 8           # 카드/패널 표준 라운드
  chip: 13          # 칩/뱃지 (pill)
  circle: 9999

# ─────────────────────────────────────────────────────────────
# COMPONENTS  (토큰 참조로 정의 — 슬라이드 생성 시 직접 인용)
# ─────────────────────────────────────────────────────────────
components:
  slide-frame:
    backgroundColor: "{colors.surface}"
    note: "16:9 1280x720 고정. 모든 콘텐츠 슬라이드는 동일 골격(kicker/chapter/title/title-rule/content/source/page)을 상속."
  kicker:
    typography: "{typography.kicker}"
    textColor: "{colors.primary}"        # 토픽 키워드는 primary, 뒤 일반어는 ink로 2색 분리
    accentTail: "{colors.ink}"
    position: "{layout.kicker}"
  chapter-indicator:
    typography: "{typography.chapter-indicator}"
    textColor: "{colors.primary-steel}"
    backgroundColor: "{colors.surface}"
    border: "1px solid {colors.divider}"
    rounded: "{rounded.md}"
    padding: "4px 10px"
    position: "{layout.chapter-indicator}"
    example: "I. 제안 개요"
  slide-title:
    typography: "{typography.slide-title}"
    textColor: "{colors.ink}"
    emphasisColor: "{colors.primary}"     # 결론 내 핵심어만 primary 강조
    position: "{layout.slide-title}"
    rule: "한 슬라이드의 결론(메시지)을 완전한 문장으로. 명사 나열 금지. (요구사항 6)"
  title-rule:
    color: "{colors.divider}"
    thickness: 1
    note: "제목과 본문 분리용 1px 헤어라인. 색 띠(color bar)·두꺼운 액센트 스트라이프 금지."
  content-card:
    backgroundColor: "{colors.surface}"
    border: "1px solid {colors.divider}"
    rounded: "{rounded.card}"
    padding: "{spacing.md}"
  panel-tint:
    backgroundColor: "{colors.tint-blue-4}"
    border: "1px solid {colors.tint-blue-3}"
    rounded: "{rounded.card}"
    padding: "{spacing.md}"
  column-header:
    backgroundColor: "{colors.primary-soft}"   # 라이트블루 면 위 잉크 텍스트
    textColor: "{colors.ink}"
    typography: "{typography.section-header}"
    rounded: "{rounded.sm}"
    padding: "8px 12px"
    align: center
  column-header-dark:
    backgroundColor: "{colors.primary-deep}"    # 다크 밴드 위 흰 텍스트 (강조 칼럼)
    textColor: "{colors.surface}"
    typography: "{typography.section-header}"
    rounded: "{rounded.sm}"
    padding: "8px 12px"
    align: center
  block-header:
    backgroundColor: "{colors.divider-soft}"
    textColor: "{colors.ink}"
    typography: "{typography.block-header}"
    rounded: "{rounded.sm}"
    padding: "6px 10px"
    align: center
  callout-chip:                       # 화살표 위 강조 칩 (예: '의사결정 강화')
    backgroundColor: "{colors.accent-coral}"
    textColor: "{colors.surface}"
    typography: "{typography.body-sm-strong}"
    rounded: "{rounded.chip}"
    padding: "3px 12px"
  bullet-item:
    marker: "■"                       # HCG 표준 사각 불릿 (primary 색)
    markerColor: "{colors.primary}"
    typography: "{typography.body}"
    textColor: "{colors.ink}"
    gap: "{spacing.xs}"
  data-table:
    headerBg: "{colors.tint-blue-1}"
    headerText: "{colors.ink}"
    headerType: "{typography.body-sm-strong}"
    cellType: "{typography.body-sm}"
    cellText: "{colors.ink}"
    rowDivider: "1px solid {colors.divider}"
    zebra: "{colors.surface-gray}"     # 짝수행 채움 (선택)
    emphasisRowBg: "{colors.tint-blue-2}"
  kpi-stat:                            # 대형 수치 콜아웃
    number: "{typography.kpi-number}"
    numberColor: "{colors.primary}"
    unit: "{typography.kpi-unit}"
    unitColor: "{colors.text-muted}"
    label: "{typography.caption}"
  process-flow:                        # 단계 화살표 흐름
    nodeBg: "{colors.tint-blue-3}"
    nodeText: "{colors.ink}"
    nodeType: "{typography.body-sm-strong}"
    arrowColor: "{colors.text-muted}"
    rounded: "{rounded.card}"
  source-note:
    typography: "{typography.source-note}"
    textColor: "{colors.text-muted}"
    position: "{layout.source-note}"
    prefix: "※ Source : "              # HCG 표준 접두 (요구사항 10)
  page-number:
    typography: "{typography.page-number}"
    textColor: "{colors.text-muted}"
    markColor: "{colors.hcg-maroon}"   # 'HCG' 워드마크
    position: "{layout.page-number}"

# ─────────────────────────────────────────────────────────────
# CHARTS  (요구사항 3: 차트/그래프 최대화)
# 정량 데이터는 표가 아닌 차트를 1순위로. 색은 아래 순서 고정.
# ─────────────────────────────────────────────────────────────
charts:
  library: "Chart.js 또는 SVG (HTML 슬라이드 기준). 폰트는 전부 Pretendard."
  series-palette:                      # 시리즈 색 순서 (범주형)
    - "{colors.primary}"               # 1: 메인 블루
    - "{colors.accent-coral}"          # 2: 코랄
    - "{colors.accent-teal}"           # 3: 틸
    - "{colors.accent-gold}"           # 4: 골드
    - "{colors.primary-soft}"          # 5: 라이트블루
    - "{colors.primary-steel}"         # 6: 딥스틸
  semantic:
    increase: "{colors.positive-green}"
    decrease: "{colors.emphasis-red}"
    highlight: "{colors.accent-coral}"
    baseline: "{colors.text-muted}"
  axis:
    color: "{colors.text-muted}"
    gridColor: "{colors.divider-soft}"
    typography: "{typography.chart-axis}"
  dataLabel:
    typography: "{typography.chart-value}"
    color: "{colors.ink}"
  types:
    - "bar / column : 항목 비교, 연도별 추이"
    - "stacked-bar : 구성비 변화(예: 보상 항목 mix)"
    - "line / area : 시계열 추세(예: 임금 경쟁력 추이)"
    - "donut / pie : 단일 구성비(범주 5개 이내)"
    - "bullet : 목표 대비 실적(KPI 달성률)"
    - "waterfall : 증감 분해(예: 인건비 변동 요인)"
    - "radar : 역량/직무 프로파일 비교"
    - "heatmap : 직급x항목 매트릭스(예: 직무평가 점수)"
---

## Overview

본 시스템은 ㈜휴먼컨설팅그룹(HCG)이 대기업 HR 컨설팅 제안서에서 실제 사용하는 시각 언어를, **4:3 → 16:9(1280×720)** 로 재설계한 슬라이드 디자인 시스템이다. 핵심 정체성은 세 가지로 요약된다. 첫째, **단일 메인 블루({colors.primary})** 를 중심축으로 라이트블루 틴트 면·코랄({colors.accent-coral}) 강조·골드/틸 보조를 더한 절제된 컨설팅 팔레트. 둘째, **11pt 주력 본문**과 8pt 출처로 대표되는 고밀도 타이포그래피 — 빈 공간을 최소화하고 한 슬라이드에 메시지·논거·시각자료를 모두 담는다. 셋째, 모든 콘텐츠 슬라이드가 동일하게 상속하는 **고정 골격(rigid skeleton)** — 좌상단 키커, 우상단 챕터 박스, 풀폭 결론형 제목, 1px 룰, 본문, 좌하단 출처, 우하단 페이지번호가 슬라이드 불문 같은 좌표·같은 크기로 고정된다.

원본 HCG deck은 도형/프로세스 다이어그램 중심이었으나, 본 시스템은 **정량 데이터를 차트로 강제**하여(요구사항 3) 설득력과 밀도를 동시에 끌어올린다. 폰트는 맑은 고딕을 **Pretendard**(메트릭 호환 오픈소스 대체 폰트)로 일원화하여 환경 간 렌더링 편차를 제거한다.

**Key Characteristics:**
- 16:9 1280×720 **전용 캔버스**, px 단위 절대 좌표 그리드 (요구사항 7, 8)
- **헤더·푸터 고정 골격**: kicker / chapter-indicator / slide-title / title-rule / source-note / page-number가 항상 동일 위치·동일 폰트 (요구사항 1, 2)
- **결론형 제목**: 명사 나열이 아닌 완결 문장으로 슬라이드 메시지를 단언 (요구사항 6)
- **Pretendard 단일 폰트** 전 요소 적용 (요구사항 4)
- **고밀도 레이아웃**: 작은 일관 간격(block-gap 12px), 본문 11pt, 콘텐츠 존을 차트/도형으로 충전 (요구사항 5, 9)
- **차트 우선**: 수치는 표보다 차트로. 6색 시리즈 팔레트 고정 (요구사항 3)
- **출처 상시 표기**: 모든 콘텐츠 슬라이드 좌하단 8pt `※ Source :` (요구사항 10)
- 메인 블루 60–70% 비중, 코랄은 강조 한정의 단일 액센트 (색 위계 dominance)

---

## 요구사항 → 구현 매핑 (Compliance Matrix)

> 사용자 지정 10개 요구사항을 본 시스템이 어떻게 하드 제약으로 강제하는지의 추적표. 슬라이드 생성·QA 시 이 표를 체크리스트로 사용한다.

| # | 요구사항 | 구현 방식 (token / rule) | 검증 포인트 |
|---|---|---|---|
| 1 | 제목·부제·챕터명·페이지수·본문 위치 고정 | `layout.kicker/chapter-indicator/slide-title/content-area/source-note/page-number` 절대 px 좌표 | 전 슬라이드 동일 좌표 오버레이 시 어긋남 0 |
| 2 | 제목·부제·챕터·페이지 폰트 크기 고정 | `typography.*` role별 단일 size/weight, 슬라이드별 재정의 금지 | role 외 임의 size 사용 0 |
| 3 | 차트·그래프 최대화 | `charts.*` 팔레트·타입 8종, "정량 데이터=차트 1순위" 규칙 | 수치 슬라이드 중 차트 포함률 ≥ 70% |
| 4 | Pretendard만 사용 | `fontPolicy: Pretendard ONLY`, 전 `typography.fontFamily=Pretendard` | 폰트 임베드/링크 Pretendard 단일 |
| 5 | 밀도 유지·빈 공간 최소화 | `content-area` 충전 규칙, `block-gap 12`, 본문 15px(11pt) | 콘텐츠 존 점유율 ≥ 85% |
| 6 | 제목=슬라이드 결론 문장 | `slide-title.rule` 완결문 강제, 명사 나열 금지 | 제목이 주어+서술어 문장인지 |
| 7 | 16:9 전용 | `canvas: 16:9` / `canvasPx: 1280x720` 고정 | 캔버스 비율 1.7778 |
| 8 | 레이아웃 불괴(견고) | px 절대 좌표 + safe-margin + maxLines 제약 | 텍스트 오버플로/요소 겹침 0 |
| 9 | 본문 빈 공간 과다 금지 | 콘텐츠 존 그리드 분할(2/3/4-up) 충전, 여백 ≤ block-gap | 단일 블록 고립·과대 여백 0 |
| 10 | 출처 표기 | `source-note` 좌하단 8pt, `prefix "※ Source : "` 상시 | 콘텐츠 슬라이드 출처 누락 0 |

---

## Layout Grid (16:9 고정 골격)

> Source: 첨부 HCG deck 8종의 slideLayout/slideMaster + 슬라이드 텍스트프레임 좌표 파싱(4:3 10.83×7.5in) → 16:9 1280×720 비례 재배치.

콘텐츠 슬라이드는 아래 7개 고정 요소를 **반드시** 상속한다. 좌표는 모두 px(좌상단 원점), 슬라이드 불문 동일.

| 영역 | 요소 | x / 위치 | y | w | 폰트 role | 비고 |
|---|---|---|---|---|---|---|
| 헤더 | kicker (토픽 라벨) | 56 (left) | 28 | 760 | kicker(18/700) | 핵심어 primary + 일반어 ink 2색 |
| 헤더 | chapter-indicator (챕터 박스) | 우측 끝 1224 정렬 | 28 | ~150 | chapter-indicator(13/600) | 1px 박스, 예 "I. 제안 개요" |
| 헤더 | slide-title (결론 제목) | 56 | 64 | 1168 | slide-title(22/700) | 완결문, 최대 2줄 |
| 헤더 | title-rule (구분 룰) | 56→1224 | 136 | 1168 | — | 1px {colors.divider} |
| 본문 | content-area | 56 | 152 | 1168 | (본문 role 혼합) | h=512, 차트/도형 충전 |
| 푸터 | source-note (출처) | 56 | 678 | 980 | source-note(11/400) | `※ Source : ...` |
| 푸터 | page-number (+HCG마크) | 우측 끝 1224 정렬 | 678 | — | page-number(13/600) | 숫자 muted + "HCG" maroon |

**Content-area 분할 규칙(고밀도용).** 콘텐츠 존(1168×512)은 다음 그리드 중 하나로 분할하고, 분할 셀을 카드/차트/도형으로 충전한다. 칼럼 간격은 `gutter 24`, 블록 간격은 `block-gap 12`.

| 패턴 | 분할 | 용도 |
|---|---|---|
| 2-up | 좌/우 572px | 비교(As-Is/To-Be, 전통 vs AX), 도식+설명 |
| 3-up | 373px × 3 | 3단계 프로세스, 3대 축, 3개 옵션 |
| 4-up | 272px × 4 | 4영역 매트릭스, 4단계 로드맵, KPI 4종 |
| 2×2 | 4분면 | 포트폴리오/우선순위 매트릭스 |
| L-rail | 좌 320 + 우 824 | 좌측 요약 레일 + 우측 본문/차트 |

**표지(cover) 골격.** 상단 아이콘 모자이크 밴드(1280×300) → 중앙 정렬 cover-title(38/800) → cover-subtitle(20/500) → cover-date(15/400) → HCG 워드마크 → 하단 저작권 고지(source-note role). 좌상단 "Strictly Confidential" 이탤릭 라벨(caption) 권장.

---

## Colors

> Source: 8개 deck `theme1.xml` 색상표는 **완전 동일**(단일 하우스 테마 확인). 사용 빈도는 전 슬라이드 합산.

### Brand Blue (메인 계열) — 화면의 60–70%
- **Primary Blue ({colors.primary})**: HCG 시그니처. 강조 텍스트, 핵심 도형, 차트 1순위. 가장 빈번(858회).
- **Primary Deep ({colors.primary-deep})**(파생): 다크 밴드/강조 칼럼 헤더/호버. 흰 텍스트와 페어.
- **Primary Steel ({colors.primary-steel})**: 2순위 헤더/도형(428회).
- **Primary Navy ({colors.primary-navy})**: 표지 모자이크 최암부, 강대비 텍스트(50회).
- **Primary Soft ({colors.primary-soft})**: 칼럼/카드 헤더 면, 차트 2순위(252회).

### Blue Tints (면·표 채움 전용)
- **{colors.tint-blue-1}**(표 헤더), **{colors.tint-blue-2}**(강조 셀), **{colors.tint-blue-3}**(박스 배경), **{colors.tint-blue-4}**(파생, 넓은 면). 텍스트가 아닌 **면적**에만 사용.

### Accent / Semantic — 단일 액센트 원칙
- **Coral ({colors.accent-coral})**: Positive 강조·화살표·하이라이트 칩. 메인 블루에 대비되는 **유일한 마케팅 액센트**(514회).
- **Gold ({colors.accent-gold})** / **Teal ({colors.accent-teal})**: 차트 3·4순위, 보조 강조 한정.
- **Emphasis Red ({colors.emphasis-red})**: 부정/리스크/감소 **수치 전용**(479회). 강조 일반용으로 남발 금지.
- **Positive Green ({colors.positive-green})**: 증가/달성 **수치 전용**.

### Neutral / Ink
- **Ink ({colors.ink})**: 본문 기본(44,940회). **Ink Slate ({colors.ink-slate})**: 차분한 대안.
- **Text Muted ({colors.text-muted})**: 캡션·출처·비활성(10,492회, 테마 lt2).
- **Divider ({colors.divider})** / **Divider Soft ({colors.divider-soft})**: 룰·표 경계.
- **Surface ({colors.surface})**: 캔버스/카드. **Surface Gray ({colors.surface-gray})**(파생): 패널 배경.

### Brand Mark
- **HCG Maroon ({colors.hcg-maroon})**: 페이지번호 옆 "HCG" 워드마크 전용(435회). 본문/도형에 사용 금지.

---

## Typography

> Source: HCG deck 실측 폰트 사이즈 분포(11pt 10,918회로 압도적 주력 / 8pt 출처 / 14·13pt 헤더) + pt→px(96dpi, ×1.333) 환산.

### Font Family
**Pretendard 단일** (orioncactus/Pretendard, SIL OFL 1.1). 맑은 고딕의 자폭·자형을 계승한 메트릭 호환 오픈소스 폰트로, 한·영·숫자 전 글리프를 커버한다. 웹/HTML 슬라이드에서는 Pretendard Variable 또는 woff2 임베드. 본문/제목/숫자/영문 라틴 전부 동일 패밀리 — **혼용 절대 금지**(요구사항 4).

### Type Scale (role = 고정 size/weight, 슬라이드별 재정의 금지)

| role | px (pt-equiv) | weight | 용도 |
|---|---|---|---|
| cover-title | 38 (28pt) | 800 | 표지 제목 |
| slide-title | 22 (16pt) | 700 | 결론형 제목(콘텐츠 슬라이드 제목) |
| kicker | 18 (13pt) | 700 | 좌상단 토픽 라벨 |
| section-header | 18 (14pt) | 700 | 칼럼/밴드 헤더 |
| block-header | 16 (12pt) | 700 | 카드 헤더 |
| **body** | **15 (11pt)** | 400 | **주력 본문** |
| body-sm | 13 (10pt) | 400 | 표·고밀도 본문 |
| chapter-indicator / page-number | 13 (10pt) | 600 | 챕터 박스 / 페이지번호 |
| caption | 12 (9pt) | 400 | 캡션·차트 축 |
| source-note | 11 (8pt) | 400 | 하단 출처 |
| kpi-number | 44 | 800 | 대형 KPI 수치 |

**Weight 운용.** 본문 400 / 강조 700의 2단 대비를 기본으로 한다. 300(Light)·900(Black)은 사용하지 않아 환경 간 위폴백을 차단한다. 줄간격은 제목 1.28, 본문 1.45, 출처 1.30으로 고정(밀도 유지).

---

## Components

### 헤더 (고정)
**`kicker`** — 좌상단 토픽 라벨. 핵심어는 {colors.primary}, 뒤따르는 일반어는 {colors.ink}로 2색 분리(예: "**HR DX** 지향점"). `{typography.kicker}`.

**`chapter-indicator`** — 우상단 챕터/목차 위치 표시. 1px {colors.divider} 박스, `{rounded.md}`, 패딩 `4px 10px`, 텍스트 {colors.primary-steel}. 예 "I. 제안 개요".

**`slide-title`** — 결론형 제목. {colors.ink} 본문에 핵심어만 {colors.primary} 강조. 최대 2줄(maxLines 2), 초과 시 폰트 축소가 아니라 **문장 압축**. (요구사항 6, 8)

**`title-rule`** — 제목/본문 분리 1px 헤어라인 {colors.divider}. 색 띠·두꺼운 스트라이프로 대체 금지.

### 콘텐츠
**`content-card`** — 흰 배경 + 1px {colors.divider} + `{rounded.card}` + 패딩 `{spacing.md}`. 기본 정보 블록.

**`panel-tint`** — {colors.tint-blue-4} 배경 + {colors.tint-blue-3} 보더. 강조 패널/요약 박스.

**`column-header` / `column-header-dark`** — 비교 칼럼 헤더. 일반 칼럼은 {colors.primary-soft} 면(잉크 텍스트), 강조 칼럼(To-Be/AX)은 {colors.primary-deep} 다크 밴드(흰 텍스트).

**`block-header`** — 카드/소블록 헤더. {colors.divider-soft} 면, 중앙 정렬.

**`callout-chip`** — 화살표 위/단계 전환 강조 칩. {colors.accent-coral} 면 + 흰 텍스트 + `{rounded.chip}`(pill). 예 "직원경험 혁신".

**`bullet-item`** — HCG 표준 사각 불릿 "■"({colors.primary}) + `{typography.body}`. 들여쓰기 2단까지, 3단부터는 카드 분할로 전환(밀도·가독 유지).

**`data-table`** — 헤더 {colors.tint-blue-1}(굵게), 셀 `{typography.body-sm}`, 행 구분 1px {colors.divider}, 강조 행 {colors.tint-blue-2}. 수치형 표는 가능하면 차트로 대체(요구사항 3).

**`kpi-stat`** — 대형 수치 콜아웃. `{typography.kpi-number}`({colors.primary}) + 단위 {colors.text-muted} + 라벨 caption. 4-up 배치 권장.

**`process-flow`** — 단계 흐름. 노드 {colors.tint-blue-3} 면(`{rounded.card}`), 화살표 {colors.text-muted}. 단계 라벨 `{typography.body-sm-strong}`.

### 푸터 (고정)
**`source-note`** — 좌하단 출처. `{typography.source-note}`({colors.text-muted}), 접두 `※ Source : `. 내부 분석은 "HCG Analysis", 외부는 출처명·연도 명기(예: "전자공시시스템 사업보고서, 2024"). (요구사항 10)

**`page-number`** — 우하단. 숫자 {colors.text-muted} + "HCG" 워드마크 {colors.hcg-maroon}. `{typography.page-number}`.

---

## Charts & Data Viz (요구사항 3)

**원칙: 정량 데이터는 표가 아닌 차트가 1순위.** 비교·추이·구성비·달성률은 즉시 차트화한다. HTML 슬라이드 기준 Chart.js 또는 SVG, 폰트는 전부 Pretendard.

**시리즈 색 순서(고정):** ① {colors.primary} → ② {colors.accent-coral} → ③ {colors.accent-teal} → ④ {colors.accent-gold} → ⑤ {colors.primary-soft} → ⑥ {colors.primary-steel}. 증감 시맨틱은 증가 {colors.positive-green} / 감소 {colors.emphasis-red} / 강조 {colors.accent-coral}.

**축·라벨:** 축선·격자 {colors.divider-soft}, 축 텍스트 `{typography.chart-axis}`, 데이터 라벨 `{typography.chart-value}`({colors.ink}). 격자선은 최소화하고 데이터 라벨을 직접 표기(밀도·가독).

| 차트 타입 | HR 컨설팅 적용 예 |
|---|---|
| bar / column | 직급별 인원, 연도별 인건비 |
| stacked-bar | 보상 항목 mix(기본급/성과급/수당) 변화 |
| line / area | 시장 임금 경쟁력(P50 대비) 추이 |
| donut / pie | 평가등급 분포(범주 5개 이내) |
| bullet | KPI 목표 대비 달성률 |
| waterfall | 인건비 변동 요인 분해 |
| radar | 직무/역량 프로파일 비교 |
| heatmap | 직급×역량 평가 매트릭스 |

---

## Density Rules (요구사항 5, 9)

1. **콘텐츠 존 점유율 ≥ 85%.** content-area(1168×512)는 분할 그리드로 충전한다. 단일 작은 블록이 넓은 여백에 고립되면 안 된다.
2. **간격은 작고 일관되게.** 칼럼 `gutter 24`, 블록 `block-gap 12`. 임의의 큰 여백(>40px) 금지.
3. **본문은 11pt(15px) 기본**으로 정보 밀도를 확보하되, 한 셀당 텍스트는 4–6줄 이내로 끊고 초과분은 카드 분할.
4. **빈 칼럼/빈 사분면 금지.** 3-up 중 2개만 차면 2-up으로 재분할한다.
5. **여백 대신 시각자료.** 남는 공간은 차트·도식·KPI 콜아웃으로 채운다(요구사항 3과 결합).
6. 단, **밀도 ≠ 혼잡.** 요소 간 최소 12px 간격과 정렬은 유지하여 레이아웃 견고성(요구사항 8)을 해치지 않는다.

---

## Title Convention (요구사항 6)

콘텐츠 슬라이드 제목은 **그 슬라이드의 결론(So-What)을 담은 완결 문장**이다. 명사구 나열("보상체계 현황")이 아니라 메시지("연공 중심 보상구조가 성과 변별력을 저해하므로 직무·성과 기반 재설계가 필요함")로 쓴다. 이는 McKinsey식 'action title / governing thought' 및 Minto Pyramid의 수평·수직 논리를 따른 것으로, 제목만 이어 읽어도 deck 전체 스토리라인이 성립해야 한다.

- 형식: [주어/대상] + [핵심 판단/서술어]. 2줄 이내.
- 핵심어 1–2개만 {colors.primary} 강조.
- 금지: 단순 라벨, 의문문, 모호한 형용사("효과적인 방안").

---

## Source / Citation Convention (요구사항 10)

모든 콘텐츠 슬라이드는 좌하단에 출처를 표기한다. `source-note` role(8pt/11px, {colors.text-muted}), 접두 `※ Source : `.

- 내부 산출물: `※ Source : HCG Analysis`
- 외부 자료: 출처명 + 매체 + 연도. 예 `※ Source : 전자공시시스템 사업보고서·기업 홈페이지, 2024`
- 데이터 차트: 데이터 출처와 기준 시점(예 `n=327, 2024.12 기준`) 병기.
- 표지: 중앙 하단에 ㈜휴먼컨설팅그룹 저작권 고지문(편집저작물·컴퓨터프로그램저작물 보호) 유지.

---

## Do's and Don'ts

### Do
- 모든 콘텐츠 슬라이드에 **고정 골격 7요소**(kicker·chapter·title·rule·content·source·page)를 상속한다.
- 메인 블루({colors.primary})를 화면의 60–70%로, 코랄은 **강조 한정 단일 액센트**로 쓴다.
- 정량 데이터는 **차트 우선**, 시리즈 색은 정해진 순서대로.
- 제목은 **결론 문장**, 핵심어만 primary 강조.
- 간격은 `gutter 24` / `block-gap 12`로 **일관**되게.
- 폰트는 전 요소 **Pretendard 단일**.

### Don't
- **색 띠/액센트 스트라이프 금지** — 제목 밑줄, 카드 한 변 컬러 보더, 슬라이드 폭 헤더/푸터 바는 AI틱하다. 구분이 필요하면 면 틴트·1px 룰·아이콘으로.
- **제목 명사 나열 금지** — 결론 없는 라벨 제목 불가(요구사항 6 위반).
- **role 외 임의 폰트 사이즈 금지** — 슬라이드별 제목/페이지 크기 변경 불가(요구사항 2 위반).
- **4:3·기타 비율 금지** — 16:9 1280×720 외 캔버스 불가(요구사항 7).
- **넓은 빈 여백 금지** — 고립된 작은 블록·빈 사분면 불가(요구사항 5, 9).
- **{colors.emphasis-red}/{colors.positive-green} 남용 금지** — 증감 수치 시맨틱 전용.
- **{colors.hcg-maroon}를 본문/도형에 사용 금지** — 워드마크 전용.
- **텍스트 오버플로/요소 겹침 금지** — 초과 시 문장 압축·셀 분할로 해결, 폰트 무단 축소 금지(요구사항 8).

---

## QA Checklist (생성 후 필수 점검)

1. **골격 일치**: 임의 2개 슬라이드를 겹쳤을 때 kicker/chapter/title/source/page 좌표가 정확히 일치하는가? (요구사항 1)
2. **폰트 사이즈 고정**: 제목 22px·챕터 13px·페이지 13px·출처 11px가 전 슬라이드 동일한가? role 외 사이즈가 있는가? (요구사항 2)
3. **차트 비중**: 수치 포함 슬라이드의 70% 이상이 차트를 쓰는가? (요구사항 3)
4. **폰트 단일**: Pretendard 외 폰트가 임베드/참조되었는가? (요구사항 4)
5. **밀도**: 콘텐츠 존 점유율 ≥ 85%? 고립 블록·과대 여백이 있는가? (요구사항 5, 9)
6. **제목 결론성**: 모든 제목이 완결 문장(주어+서술어)인가? 명사 나열은 없는가? (요구사항 6)
7. **비율**: 캔버스가 1280×720(16:9)인가? (요구사항 7)
8. **견고성**: 텍스트 오버플로·요소 겹침·정렬 이탈이 0인가? (요구사항 8)
9. **출처**: 모든 콘텐츠 슬라이드 좌하단에 `※ Source :` 가 있는가? (요구사항 10)
10. **색 위계**: 메인 블루 dominance 유지, 색 띠/밑줄/스트라이프 미사용?

---

## Known Gaps

- **HCG 로고 마룬({colors.hcg-maroon})** 은 deck XML 최빈 다크레드(#921F0B)로 지정했으나, 정식 BI 매뉴얼의 Pantone/HEX 원값으로 교체 권장.
- **표지 아이콘 모자이크 밴드**는 프로젝트별 커스텀 일러스트 영역으로, 본 시스템은 그리드(1280×300)·색면 규칙만 정의한다. 아이콘 세트는 단색 라인(stroke {colors.primary}) 통일 권장.
- **애니메이션/전환**은 정적 deck 기준이라 미정의. HTML 슬라이드 적용 시 150–250ms ease-out 권장.
- **다크 모드**는 컨설팅 제안서 특성상 미정의(흰 캔버스 고정).
- 원본 deck은 **4:3**이므로, 기존 자산을 16:9로 옮길 때 가로 여유 공간(+2.5in 상당)은 차트 확대·L-rail 추가로 활용 권장.

---

## Sources

1. **HCG 디자인 토큰 일체** — 첨부 HCG 제안서 deck 8종(_HCG_리파인/DB하이텍/HJ_NEW/SK디앤디/롯데케미칼/일진/포스코인터내셔널/기아)의 `theme1.xml`·`slideMaster`·`slideLayout`·slide XML 직접 파싱. 색상표 8종 완전 일치(단일 하우스 테마) 확인.
2. **pt→px 환산(1pt = 1.333px @96dpi)** — W3C CSS Values and Units(절대 길이 단위 1in=96px, 1pt=1/72in).
3. **16:9 = 1280×720** — 표준 와이드스크린 프레젠테이션/웹 캔버스 규격(13.333×7.5in @96dpi).
4. **Pretendard** — orioncactus/Pretendard, SIL Open Font License 1.1.
5. **명도대비 4.5:1(본문)** — W3C WCAG 2.1 §1.4.3 Contrast (Minimum), AA.
6. **결론형 제목(action title / governing thought) 및 수평·수직 논리** — Barbara Minto, *The Pyramid Principle*; McKinsey 스토리라인 작성 원칙.
