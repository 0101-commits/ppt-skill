---
name: hcg-ppt-design
description: >
  HCG 제안서/발표 슬라이드의 시각 언어 정본(v2.0). 16:9 1280×720 px 고정 캔버스,
  Pretendard 단일 폰트, 단일 메인 블루(#356CB5) 팔레트, 고정 골격(rigid skeleton),
  결론형 제목, 차트 우선, 고밀도(≥85%) 레이아웃을 정의한다.
  구(舊) 4:3 / 맑은 고딕 / HCG_RED(#921F0B) / 템플릿 상속 시스템을 전면 폐기·대체한다.
when_to_use: >
  확정된 목차/장표(skill_ppt_planning 산출물)를 실제 슬라이드로 구현할 때.
  색상·폰트·좌표·컴포넌트 토큰을 적용하고, core/designer.py(PPTX) 또는
  core/html_renderer.py(HTML)로 자동 생성할 때.
---

# HCG 슬라이드 디자인(Design) 스킬 v2.0

> 절대 기준: `HCG-slide-design-system.md` (HCG-Slide-Design-System v1.0). 기계 판독 페어: `skill_ppt_design.json`. 본 문서는 그 표준을 사람이 읽는 정본(正本)으로 미러링한 것이며, 토큰명·색상·role명은 JSON 계약과 정확히 일치한다.

---

## Overview

본 시스템은 ㈜휴먼컨설팅그룹(HCG)이 대기업 HR 컨설팅 제안서에서 실제 사용하는 시각 언어를, **4:3 → 16:9(1280×720 px)** 로 전면 재설계한 슬라이드 디자인 시스템이다. 핵심 정체성은 세 가지로 요약된다.

1. **단일 메인 블루(`primary` #356CB5)** 를 중심축으로, 라이트블루 틴트 면·코랄(`accent-coral` #F16249) 강조·골드/틸 보조를 더한 절제된 컨설팅 그레이드 팔레트.
2. **11pt 주력 본문**(`body` 15px)과 8pt 출처(`source-note` 11px)로 대표되는 고밀도 타이포그래피 — 빈 공간을 최소화하고 한 슬라이드에 메시지·논거·시각자료를 모두 담는다.
3. 모든 콘텐츠 슬라이드가 동일하게 상속하는 **고정 골격(rigid skeleton)** — 좌상단 키커, 우상단 챕터 박스, 풀폭 결론형 제목, 1px 룰, 본문, 좌하단 출처, 우하단 페이지번호가 슬라이드 불문 같은 좌표·같은 크기로 고정된다.

원본 HCG deck은 도형/프로세스 다이어그램 중심이었으나, 본 시스템은 **정량 데이터를 차트로 강제**(요구사항 3)하여 설득력과 밀도를 동시에 끌어올린다. 폰트는 맑은 고딕을 **Pretendard**(메트릭 호환 오픈소스 대체 폰트)로 일원화하여 환경 간 렌더링 편차를 제거한다.

### v1.5 → v2.0 변경 (전면 재설계)

| 항목 | 구(v1.5, 폐기) | 신(v2.0) |
|---|---|---|
| 캔버스 | 4:3 (10.833×7.5 in) | **16:9 (1280×720 px)** |
| 단위 | inch | **px** (좌상단 원점, `px→in = /96`, `pt→px = ×1.333`) |
| 폰트 | 맑은 고딕 | **Pretendard 단일** |
| 메인 색 | HCG_RED `#921F0B` | **BLUE `#356CB5`** (마룬은 워드마크 전용으로 강등) |
| 생성 방식 | 템플릿 상속 / 테마(hcg·paradise) | **빈 16:9 캔버스에서 생성**, 고정 골격 + 토큰 주입 |
| 엔진 | 단일 PPTX | **2종**: `core/designer.py`(PPTX) + `core/html_renderer.py`(HTML) |

**Key Characteristics**
- 16:9 1280×720 **전용 캔버스**, px 단위 절대 좌표 그리드 (요구사항 7, 8)
- **헤더·푸터 고정 골격**: `kicker` / `chapter-indicator` / `slide-title` / `title-rule` / `source-note` / `page-number`가 항상 동일 위치·동일 폰트 (요구사항 1, 2)
- **결론형 제목**: 명사 나열이 아닌 완결 문장으로 슬라이드 메시지를 단언 (요구사항 6)
- **Pretendard 단일 폰트** 전 요소 적용 (요구사항 4)
- **고밀도 레이아웃**: 작은 일관 간격(`block-gap` 12px), 본문 11pt, 콘텐츠 존을 차트/도형으로 충전 (요구사항 5, 9)
- **차트 우선**: 수치는 표보다 차트로. 6색 시리즈 팔레트 고정 (요구사항 3)
- **출처 상시 표기**: 모든 콘텐츠 슬라이드 좌하단 8pt `※ Source :` (요구사항 10)
- 메인 블루 60–70% 비중, 코랄은 강조 한정의 단일 액센트 (색 위계 dominance)

---

## 요구사항 → 구현 매핑 (Compliance Matrix)

> 사용자 지정 10개 요구사항을 본 시스템이 어떻게 하드 제약으로 강제하는지의 추적표. 슬라이드 생성·QA 시 이 표를 체크리스트로 사용한다.

| # | 요구사항 | 구현 방식 (token / rule) | 검증 포인트 |
|---|---|---|---|
| 1 | 제목·부제·챕터명·페이지수·본문 위치 고정 | `layout.kicker/chapter-indicator/slide-title/content-area/source-note/page-number` 절대 px 좌표 | 전 슬라이드 동일 좌표 오버레이 시 어긋남 0 |
| 2 | 제목·부제·챕터·페이지 폰트 크기 고정 | `typography.*` role별 단일 px/weight, 슬라이드별 재정의 금지 | role 외 임의 size 사용 0 |
| 3 | 차트·그래프 최대화 | `charts.*` 팔레트·타입 8종+, "정량 데이터=차트 1순위" 규칙 | 수치 슬라이드 중 차트 포함률 ≥ 70% |
| 4 | Pretendard만 사용 | `font_policy: Pretendard ONLY`, 전 `typography.fontFamily=Pretendard` | 폰트 임베드/링크 Pretendard 단일 |
| 5 | 밀도 유지·빈 공간 최소화 | `content-area` 충전 규칙, `block-gap 12`, 본문 15px(11pt) | 콘텐츠 존 점유율 ≥ 85% |
| 6 | 제목=슬라이드 결론 문장 | `slide-title.rule` 완결문 강제, 명사 나열 금지 | 제목이 주어+서술어 문장인지 |
| 7 | 16:9 전용 | `canvas: 16:9` / `1280×720` 고정 | 캔버스 비율 1.7778 |
| 8 | 레이아웃 불괴(견고) | px 절대 좌표 + `safe-margin` + `maxLines` 제약 | 텍스트 오버플로/요소 겹침 0 |
| 9 | 본문 빈 공간 과다 금지 | 콘텐츠 존 그리드 분할(2/3/4-up) 충전, 여백 ≤ `block-gap` | 단일 블록 고립·과대 여백 0 |
| 10 | 출처 표기 | `source-note` 좌하단 8pt, `prefix "※ Source : "` 상시 | 콘텐츠 슬라이드 출처 누락 0 |

---

## Canvas & Font Policy

### Canvas (요구사항 7, 8)

| 속성 | 값 |
|---|---|
| 비율 | **16:9 전용** (1.7778) — 4:3·기타 비율 금지 |
| 픽셀 | **1280 × 720 px** |
| 인치 | 13.333 × 7.5 in @ 96dpi |
| 단위 | **px**, 좌상단 원점 |
| 환산 | `pt → px = pt × 1.333` (96dpi) · `px → in = px / 96` |
| PPTX 변환 | `PX() = Inches(px / 96)` |
| 배경 | `surface` #FFFFFF 고정 (다크모드 미정의) |

> HCG 원본 deck의 pt 실측값을 px로 환산해 고정한다(11pt 주력 본문 → 15px). 모든 좌표·크기는 px 절대값이며 슬라이드 불문 동일하다.

### Font Policy (요구사항 4)

**Pretendard 단일** (orioncactus/Pretendard, SIL OFL 1.1). 맑은 고딕의 자폭·자형을 계승한 메트릭 호환 오픈소스 폰트로, 한·영·숫자 전 글리프를 커버한다. 본문/제목/숫자/영문 라틴 **전부 Pretendard 단일 — 혼용 절대 금지.**

- **사용 weight**: `400 / 500 / 600 / 700 / 800` (본문 400 · 강조 700 2단 대비 기본)
- **금지 weight**: `300`(Light) · `900`(Black) — 환경 간 위(僞)폴백 차단
- **PPTX**: run XML `a:latin` + `a:ea` `typeface='Pretendard'`. 렌더링을 위해 시스템 설치 또는 폰트 임베드 필요.
- **HTML**: Pretendard woff2/Variable `@font-face` 또는 CDN. 폴백 없이 Pretendard 단일.

---

## Colors

> Source: 8개 deck `theme1.xml` 색상표는 **완전 동일**(단일 하우스 테마 확인). 사용 빈도는 전 슬라이드 합산. 색상값·토큰명은 JSON 계약과 정확히 일치한다.

### Brand Blue (메인 계열) — 화면의 60–70%

| 토큰 | HEX | 용도 |
|---|---|---|
| `primary` | **#356CB5** | HCG 시그니처. 강조 텍스트·핵심 도형·차트 1순위 (최빈 858회) |
| `primary-deep` | #1F4E8C | (파생) 다크 밴드/강조 칼럼 헤더/호버. 흰 텍스트와 페어 |
| `primary-steel` | #336699 | 2순위 헤더/도형 (428회) |
| `primary-navy` | #0D1467 | 표지 모자이크 최암부·강대비 텍스트 (50회) |
| `primary-soft` | #A1D1F1 | 칼럼/카드 헤더 면·차트 2순위 (252회) |

### Blue Tints (면·표 채움 전용)

`tint-blue-1` #D4E1F2(표 헤더), `tint-blue-2` #B0CEF2(강조 셀), `tint-blue-3` #A5C4E1(박스 배경), `tint-blue-4` #EAF2FB(넓은 면). 텍스트가 아닌 **면적**에만 사용.

### Accent / Semantic — 단일 액센트 원칙

| 토큰 | HEX | 용도 |
|---|---|---|
| `accent-coral` | #F16249 | Positive 강조·화살표·하이라이트 칩. 메인 블루에 대비되는 **유일한 마케팅 액센트** (514회) |
| `accent-gold` | #FFCC66 | 차트 4순위·보조 강조 한정 |
| `accent-teal` | #50B8B6 | 차트 3순위·보조 강조 한정 |
| `emphasis-red` | #C00000 | 부정/리스크/감소 **수치 전용** (479회). 일반 강조 남발 금지 |
| `positive-green` | #00B050 | 증가/달성 **수치 전용** |

### Neutral / Ink

`ink` #000000(본문 기본, 44,940회), `ink-slate` #263238(차분한 대안), `text-muted` #919191(캡션·출처·비활성, 테마 lt2), `text-faint` #B7B7B7(플레이스홀더), `divider` #DDDDDD / `divider-soft` #EAEAEA(룰·표 경계), `surface` #FFFFFF(캔버스/카드), `surface-gray` #F5F6F8(패널 배경).

### Brand Mark

`hcg-maroon` **#921F0B** — 페이지번호 옆 "HCG" 워드마크 **전용** (435회). **본문/도형에 사용 금지.**

### Color Dominance Rule (색 위계)

- **메인 블루 `primary`가 화면의 60~70%.** 코랄 `accent-coral`은 **강조 한정 단일 액센트**.
- `tint-blue-*`는 텍스트가 아닌 **면적(표/박스 채움)에만**.
- **시맨틱 전용**: `emphasis-red`=부정/감소 수치, `positive-green`=증가/달성 수치. 일반 강조 남용 금지.
- `hcg-maroon`은 'HCG' 워드마크 전용. 본문/도형 금지.
- **색 띠/액센트 스트라이프 금지** — 제목 밑줄 컬러바·카드 한 변 컬러 보더·슬라이드폭 헤더/푸터 바는 사용하지 않는다. 구분이 필요하면 **면 틴트·1px 룰·아이콘**으로.

---

## Typography

> Source: HCG deck 실측 폰트 사이즈 분포(11pt 압도적 주력 / 8pt 출처 / 14·13pt 헤더) + pt→px(96dpi, ×1.333) 환산. **각 role의 px/weight는 슬라이드 불문 고정값. 슬라이드별 재정의 금지** (요구사항 2).

| role | px (pt-equiv) | weight | lineHeight | 용도 |
|---|---|---|---|---|
| `cover-title` | 38 (28pt) | 800 | 1.22 | 표지 제목 |
| `cover-subtitle` | 20 (15pt) | 500 | 1.40 | 표지 부제 |
| `cover-date` | 15 | 400 | 1.40 | 표지 일자 |
| `slide-title` | 22 (16pt) | 700 | 1.28 | 결론형 제목(콘텐츠 슬라이드 제목), 최대 2줄 |
| `kicker` | 18 (13pt) | 700 | 1.20 | 좌상단 토픽 라벨 |
| `chapter-indicator` | 13 (10pt) | 600 | 1.10 | 우상단 챕터 박스 |
| `section-header` | 18 (14pt) | 700 | 1.25 | 칼럼/밴드 헤더 |
| `block-header` | 16 (12pt) | 700 | 1.25 | 카드 헤더 |
| **`body`** | **15 (11pt)** | 400 | 1.45 | **주력 본문** |
| `body-strong` | 15 | 700 | 1.45 | 본문 강조 |
| `body-sm` | 13 (10pt) | 400 | 1.40 | 표·고밀도 본문 |
| `body-sm-strong` | 13 | 700 | 1.40 | 표·고밀도 본문 강조 |
| `caption` | 12 (9pt) | 400 | 1.35 | 캡션 |
| `source-note` | 11 (8pt) | 400 | 1.30 | 하단 출처 |
| `page-number` | 13 (10pt) | 600 | 1.10 | 우하단 페이지번호 |
| `chart-axis` | 12 | 500 | 1.20 | 차트 축/범례 |
| `chart-value` | 13 | 700 | 1.10 | 차트 데이터 라벨 |
| `kpi-number` | 44 | 800 | 1.05 | 대형 KPI 수치 |
| `kpi-unit` | 16 | 600 | 1.10 | KPI 단위/라벨 |

**Weight 운용.** 본문 400 / 강조 700의 2단 대비를 기본으로 한다. 300·900은 사용하지 않는다. 줄간격은 제목 1.28, 본문 1.45, 출처 1.30으로 고정(밀도 유지).

---

## Layout Grid (16:9 고정 골격)

> Source: HCG deck 8종 slideLayout/slideMaster + 텍스트프레임 좌표 파싱(4:3) → 16:9 1280×720 비례 재배치. 좌표는 모두 px(좌상단 원점), 슬라이드 불문 동일 (요구사항 1, 7, 8).

### 캔버스 기본

- `canvas`: w 1280, h 720
- `safe-margin`: left 56, right 56, top 28, bottom 28 → content box x:[56, 1224], y:[28, 692]
- `content-width`: 1168 (= 1280 − 56×2)

### 고정 골격 7요소 (rigid skeleton)

콘텐츠 슬라이드는 아래 7개 고정 요소를 **반드시** 상속한다. 엔진이 자동 적용한다.

| 영역 | 요소(token) | x / 위치 | y | w | 폰트 role | 비고 |
|---|---|---|---|---|---|---|
| 헤더 | `kicker` | 56 (left) | 28 | 760 | kicker(18/700) | 핵심어 `primary` + 일반어 `ink` 2색 |
| 헤더 | `chapter-indicator` | 우측 끝 1224 정렬 | 28 | ~150 | chapter-indicator(13/600) | 1px 박스, 예 "I. 제안 개요" |
| 헤더 | `slide-title` | 56 | 64 | 1168 | slide-title(22/700) | 완결문, 최대 2줄 |
| 헤더 | `title-rule` | 56 → 1224 | 136 | 1168 | — | 1px `divider` |
| 본문 | `content-area` | 56 | 152 | 1168 (h=512) | (본문 role 혼합) | 차트/도형 충전 |
| 푸터 | `source-note` | 56 | 678 | 980 | source-note(11/400) | `※ Source : ...` |
| 푸터 | `page-number` | 우측 끝 1224 정렬 | 678 | — | page-number(13/600) | 숫자 muted + "HCG" maroon |

### Content-area 분할 그리드 (고밀도용)

콘텐츠 존(1168×512)은 다음 그리드 중 하나로 분할하고, 분할 셀을 카드/차트/도형으로 충전한다. 칼럼 간격은 `gutter 24`, 블록 간격은 `block-gap 12`.

| 패턴 | 분할 | 용도 |
|---|---|---|
| `2-up` | 좌/우 572px × 2 | 비교(As-Is/To-Be, 전통 vs AX), 도식+설명 |
| `3-up` | 373px × 3 | 3단계 프로세스, 3대 축, 3개 옵션 |
| `4-up` | 272px × 4 | 4영역 매트릭스, 4단계 로드맵, KPI 4종 |
| `2x2` | 4분면 | 포트폴리오/우선순위 매트릭스 |
| `L-rail` | 좌 320 + 우 824 | 좌측 요약 레일 + 우측 본문/차트 |

### 표지(cover) 골격

상단 아이콘 모자이크 밴드 `cover-hero-band`(x0 y0 1280×300) → 중앙 정렬 `cover-title`(x160 y392, 38/800) → `cover-subtitle`(y470, 20/500) → `cover-date`(y540, 15/400) → `cover-logo` HCG 워드마크(x560 y612, 160×48) → 하단 저작권 고지 `cover-legal`(y676, source-note role). 좌상단 "Strictly Confidential" 이탤릭 라벨(caption) 권장.

---

## Spacing / Rounded

> 고밀도 유지: 간격은 작고 일관되게 (요구사항 5, 9).

**`spacing`** (px): `xxs` 4 · `xs` 8 · `sm` 12 · `md` 16 · `lg` 20 · `xl` 24 · `gutter` 24(칼럼 사이 표준) · `block-gap` 12(카드/블록 사이 표준, 고밀도).

**`rounded`** (px): `none` 0 · `sm` 3 · `md` 6 · `lg` 10 · `card` 8(카드/패널 표준) · `chip` 13(칩/뱃지 pill) · `circle` 9999.

---

## Components

> 토큰 참조로 정의 — 슬라이드 생성 시 직접 인용. 색상·typography·position은 모두 위 토큰을 가리킨다.

### 헤더 (고정)

- **`kicker`** — 좌상단 토픽 라벨. 핵심어는 `primary`, 뒤따르는 일반어는 `ink`로 2색 분리(예: "**HR DX** 지향점"). typography `kicker`, position `layout.kicker`.
- **`chapter-indicator`** — 우상단 챕터/목차 위치 표시. 텍스트 `primary-steel`, 배경 `surface`, `1px solid divider` 박스, `rounded md`, 패딩 `4px 10px`. 예 "I. 제안 개요".
- **`slide-title`** — 결론형 제목. `ink` 본문에 핵심어 1~2개만 `primary`(emphasisColor) 강조. 최대 2줄(maxLines 2), 초과 시 폰트 축소가 아니라 **문장 압축**. 명사 나열 금지 (요구사항 6, 8).
- **`title-rule`** — 제목/본문 분리 1px 헤어라인 `divider`. 색 띠·두꺼운 스트라이프로 대체 금지.

### 콘텐츠

- **`content-card`** — 배경 `surface` + `1px solid divider` + `rounded card` + 패딩 `spacing.md`. 기본 정보 블록.
- **`panel-tint`** — 배경 `tint-blue-4` + `1px solid tint-blue-3` 보더 + `rounded card` + 패딩 `spacing.md`. 강조 패널/요약 박스.
- **`column-header`** — 일반 비교 칼럼 헤더. 면 `primary-soft`, 텍스트 `ink`, typography `section-header`, `rounded sm`, 패딩 `8px 12px`, 중앙 정렬.
- **`column-header-dark`** — 강조 칼럼(To-Be/AX) 헤더. 면 `primary-deep`, 텍스트 `surface`(흰색), 나머지 동일.
- **`block-header`** — 카드/소블록 헤더. 면 `divider-soft`, 텍스트 `ink`, typography `block-header`, `rounded sm`, 패딩 `6px 10px`, 중앙 정렬.
- **`callout-chip`** — 화살표 위/단계 전환 강조 칩(pill). 면 `accent-coral` + 텍스트 `surface` + typography `body-sm-strong` + `rounded chip` + 패딩 `3px 12px`. 예 "직원경험 혁신".
- **`bullet-item`** — HCG 표준 사각 불릿 `■`(markerColor `primary`) + typography `body`(텍스트 `ink`) + gap `xs`. 들여쓰기 2단까지, 3단부터는 카드 분할로 전환(밀도·가독 유지).
- **`data-table`** — 헤더 배경 `tint-blue-1`/텍스트 `ink`(`body-sm-strong`), 셀 `body-sm`(텍스트 `ink`), 행 구분 `1px solid divider`, 짝수행 `surface-gray`(zebra, 선택), 강조 행 `tint-blue-2`(emphasisRowBg). 수치형 표는 가능하면 차트로 대체 (요구사항 3).
- **`kpi-stat`** — 대형 수치 콜아웃. number `kpi-number`(색 `primary`) + unit `kpi-unit`(색 `text-muted`) + label `caption`. 4-up 배치 권장.
- **`process-flow`** — 단계 흐름. 노드 면 `tint-blue-3`(텍스트 `ink`, `rounded card`), 라벨 `body-sm-strong`, 화살표 `text-muted`.

### 푸터 (고정)

- **`source-note`** — 좌하단 출처. typography `source-note`(색 `text-muted`), 접두 `※ Source : `. 내부 분석은 "HCG Analysis", 외부는 출처명·연도 명기 (요구사항 10).
- **`page-number`** — 우하단. 숫자 `text-muted` + "HCG" 워드마크 `hcg-maroon`(markColor). typography `page-number`.

---

## Charts & Data Viz (요구사항 3)

**원칙: 정량 데이터는 표가 아닌 차트가 1순위.** 비교·추이·구성비·달성률은 즉시 차트화한다. 수치 포함 슬라이드의 **차트 포함률 ≥ 70%**.

- **라이브러리**: HTML — Chart.js 또는 SVG(폰트 Pretendard). PPTX — python-pptx native chart(`CategoryChartData`, 폰트 Pretendard).
- **시리즈 색 순서(고정)**: ① `primary` → ② `accent-coral` → ③ `accent-teal` → ④ `accent-gold` → ⑤ `primary-soft` → ⑥ `primary-steel`.
- **시맨틱**: 증가 `positive-green` / 감소 `emphasis-red` / 강조 `highlight` `accent-coral` / 기준선 `baseline` `text-muted`.
- **축·라벨**: 축선·격자 `divider-soft`, 축 텍스트 `chart-axis`, 데이터 라벨 `chart-value`(색 `ink`). 격자선은 최소화하고 데이터 라벨을 직접 표기(밀도·가독).

| 차트 타입 | HR 컨설팅 적용 예 |
|---|---|
| `bar` / `column` | 직급별 인원, 연도별 인건비 |
| `stacked-bar` | 보상 항목 mix(기본급/성과급/수당) 변화 |
| `line` / `area` | 시장 임금 경쟁력(P50 대비) 추이 |
| `donut` / `pie` | 평가등급 분포(범주 5개 이내) |
| `bullet` | KPI 목표 대비 달성률 |
| `waterfall` | 인건비 변동 요인 분해 |
| `radar` | 직무/역량 프로파일 비교 |
| `heatmap` | 직급×역량 평가 매트릭스 |

---

## Slide Types

> 13개 타입(JSON `slide_types`). **모든 콘텐츠 타입은 고정 골격을 상속**하며, 공통 선택 필드 `kicker?` / `chapter?` / `source?`를 받아 골격에 채운다. `title`은 결론 문장(필수). 엔진이 타입을 components/layout 토큰에 매핑한다.

| 타입 | required | optional | 1줄 목적 |
|---|---|---|---|
| `cover` | title | subtitle, date, confidential, legal, hero | 표지: hero-band + 중앙 정렬 title/subtitle/date + HCG 워드마크 + 하단 저작권 |
| `toc` | items | current, kicker, chapter, source | 목차/섹션 디바이더. `current`=강조 인덱스, 로마숫자 칩 |
| `section` | title | num, sub | 섹션 간지(풀비주얼). `num` 자동 로마숫자 |
| `bullets` | title, items | kicker, chapter, source | `■` 불릿 본문. items=[str] 또는 [{text, sub:[..]}], 2단까지 |
| `cards` | title, items | kicker, chapter, source | content-card 그리드(2/3/4-up). items=[{header, body\|bullets}] |
| `columns` | title, columns | kicker, chapter, source | 칼럼 그리드 + column-header(/-dark). columns=[{header, dark?, body\|bullets}] |
| `compare` | title, left, right | left_header, right_header, kicker, chapter, source, takeaway | 2-up As-Is/To-Be. To-Be 헤더는 column-header-dark |
| `kpi` | title, items | kicker, chapter, source | kpi-stat 행(4-up). items=[{number, unit?, label}] |
| `process` | title, steps | kicker, chapter, source, desc | process-flow 노드 + 화살표. steps=[{label, desc?}] |
| `matrix` | title, columns, rows | kicker, chapter, source, takeaway | 차별화 매트릭스. columns=[헤더], rows=[{group, cells:[..]}] |
| `chart` | title, chart | kicker, chapter, source | 차트 슬라이드. chart={type, labels:[..], series:[{name, data:[..]}], unit?} |
| `table` | title, headers, rows | kicker, chapter, source, emphasis_row | data-table. `emphasis_row`=강조 행 인덱스 |
| `end` | — | message | 마무리 슬라이드 |

---

## Density Rules (요구사항 5, 9)

1. **콘텐츠 존 점유율 ≥ 85%.** `content-area`(1168×512)는 분할 그리드로 충전한다. 단일 작은 블록이 넓은 여백에 고립되면 안 된다.
2. **간격은 작고 일관되게.** 칼럼 `gutter 24`, 블록 `block-gap 12`. 임의의 큰 여백(>40px) 금지.
3. **본문은 11pt(15px) 기본**으로 정보 밀도를 확보하되, 한 셀당 텍스트는 4–6줄 이내로 끊고 초과분은 카드 분할.
4. **빈 칼럼/빈 사분면 금지.** 3-up 중 2개만 차면 2-up으로 재분할한다.
5. **여백 대신 시각자료.** 남는 공간은 차트·도식·KPI 콜아웃으로 채운다(요구사항 3과 결합).
6. 단, **밀도 ≠ 혼잡.** 요소 간 최소 12px 간격과 정렬은 유지하여 레이아웃 견고성(요구사항 8)을 해치지 않는다.

---

## Title Convention (요구사항 6)

콘텐츠 슬라이드 제목은 **그 슬라이드의 결론(So-What)을 담은 완결 문장**이다. 명사구 나열("보상체계 현황")이 아니라 메시지("연공 중심 보상구조가 성과 변별력을 저해하므로 직무·성과 기반 재설계가 필요함")로 쓴다. 이는 McKinsey식 'action title / governing thought' 및 Minto Pyramid의 수평·수직 논리를 따른 것으로, 제목만 이어 읽어도 deck 전체 스토리라인이 성립해야 한다.

- **형식**: [주어/대상] + [핵심 판단/서술어]. 2줄 이내. 핵심어 1–2개만 `primary` 강조.
- **금지**: 명사구 나열, 단순 라벨, 의문문, 모호한 형용사("효과적인 방안").

---

## Source / Citation Convention (요구사항 10)

모든 콘텐츠 슬라이드는 좌하단에 출처를 표기한다. `source-note` role(8pt/11px, `text-muted`), 접두 `※ Source : `.

- **내부 산출물**: `※ Source : HCG Analysis`
- **외부 자료**: 출처명 + 매체 + 연도. 예 `※ Source : 전자공시시스템 사업보고서·기업 홈페이지, 2024`
- **데이터 차트**: 데이터 출처와 기준 시점 병기. 예 `n=327, 2024.12 기준`
- **표지**: 중앙 하단에 ㈜휴먼컨설팅그룹 저작권 고지문(편집저작물·컴퓨터프로그램저작물 보호) 유지.

---

## Do's and Don'ts

### Do
- 모든 콘텐츠 슬라이드에 **고정 골격 7요소**(kicker·chapter·title·rule·content·source·page)를 상속한다.
- 메인 블루 `primary`를 화면의 60–70%로, 코랄은 **강조 한정 단일 액센트**로 쓴다.
- 정량 데이터는 **차트 우선**, 시리즈 색은 정해진 순서대로.
- 제목은 **결론 문장**, 핵심어만 `primary` 강조.
- 간격은 `gutter 24` / `block-gap 12`로 **일관**되게.
- 폰트는 전 요소 **Pretendard 단일**.

### Don't
- **색 띠/액센트 스트라이프 금지** — 제목 밑줄, 카드 한 변 컬러 보더, 슬라이드 폭 헤더/푸터 바는 AI틱하다. 구분이 필요하면 면 틴트·1px 룰·아이콘으로.
- **제목 명사 나열 금지** — 결론 없는 라벨 제목 불가 (요구사항 6).
- **role 외 임의 폰트 사이즈 금지** — 슬라이드별 제목/페이지 크기 변경 불가 (요구사항 2).
- **4:3·기타 비율 금지** — 16:9 1280×720 외 캔버스 불가 (요구사항 7).
- **넓은 빈 여백 금지** — 고립된 작은 블록·빈 사분면 불가 (요구사항 5, 9).
- **`emphasis-red`/`positive-green` 남용 금지** — 증감 수치 시맨틱 전용.
- **`hcg-maroon`를 본문/도형에 사용 금지** — 워드마크 전용.
- **텍스트 오버플로/요소 겹침 금지** — 초과 시 문장 압축·셀 분할로 해결, 폰트 무단 축소 금지 (요구사항 8).

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

## Engine (구현 엔진 2종)

본 디자인 시스템은 동일한 spec/토큰을 소비하는 **두 개의 렌더링 엔진**으로 구현된다. 양 엔진 모두 본 JSON 계약의 `colors`/`typography`/`layout`/`components`/`charts` 토큰을 소비하며, `apply_design_tokens(json)`으로 팔레트를 주입한다.

### `core/designer.py` — PPTX 엔진 (python-pptx)
- `DeckEngine` / `Designer`. **빈 16:9 캔버스**를 생성한다(`slide_width`/`slide_height = PX`). **템플릿 상속 폐기.**
- 좌표 변환: `PX() = Inches(px / 96)` (px → EMU).
- 폰트: run XML `a:latin` + `a:ea` `typeface='Pretendard'`.
- 흐름: `render(spec).save()`.

### `core/html_renderer.py` — HTML 엔진 (HTML/CSS + Chart.js)
- `HtmlRenderer`. `1280×720` `.slide` div, Pretendard `@font-face`, 차트는 Chart.js.
- 흐름: `render(spec)` → `.html` 산출.

### CLI
```
python main.py --client <name> [--out path] [--html | --pptx | --both]
```
기본값은 `--both`(PPTX + HTML 동시 생성).

---

## Known Gaps

- **`hcg-maroon` #921F0B**: deck XML 최빈 다크레드로 지정. 정식 BI 매뉴얼의 Pantone/HEX 원값으로 교체 권장.
- **표지 아이콘 모자이크 밴드(`cover-hero-band`)**: 프로젝트별 커스텀 일러스트 영역. 본 시스템은 그리드(1280×300)·색면 규칙만 정의한다. 아이콘 세트는 단색 라인(stroke `primary`) 통일 권장.
- **애니메이션/전환**: 정적 deck 기준이라 미정의. HTML 슬라이드 적용 시 150–250ms ease-out 권장.
- **다크 모드**: 컨설팅 제안서 특성상 미정의(흰 캔버스 고정).
- **PPTX + Pretendard**: PPTX 렌더링은 Pretendard **시스템 설치 또는 폰트 임베드** 전제. 미설치 시 PowerPoint 폴백 폰트로 표시된다.
- **PPTX 고급 차트**: `waterfall` / `radar` / `heatmap` / `bullet`은 python-pptx 네이티브 미지원 → **HTML(Chart.js) 우선** 또는 도형 근사로 처리.

---

## Handoff (from skill_ppt_planning)

- **From**: `skill_ppt_planning.json` — 스토리/메시지/목차를 결정하는 기획 페어.
- **입력**: 확정 목차 + 장표별 결론 제목 + 슬라이드 타입 태그(`slide_types`) + (선택) `kicker`/`chapter`/`source`.
- **출력**: 타입을 `components`/`layout` 토큰에 매핑해 PPTX/HTML 슬라이드 생성.
- 즉, **기획(skill_ppt_planning)이 "무엇을 말할지"를 정하고, 본 스킬(hcg-ppt-design)이 "어떻게 보일지"를 고정 골격·토큰으로 구현한다.**

---

v2.0 — HCG-Slide-Design-System v1.0 절대 기준 전면 재설계 | JSON 페어: skill_ppt_design.json | 기획 페어: skill_ppt_planning
