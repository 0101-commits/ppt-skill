---
name: hcg-ppt-design
description: >
  HCG 제안서 PPT 디자인 & 자동생성 가이드라인 v6.0.
  기아 _final + 롯데알미늄 _final PPTX 전수 역엔지니어링 — 구조/레이아웃/색상/좌표 정합 +
  Text-to-Visual 구조화 엔진(목차 어젠다카드 / N등분 카드 그리드 / 컨테이너) +
  XML 레벨 고급 디자인(윤곽선/그림자/투명도/그라디언트/자간) + DeckEngine 통합 엔진.
  신규 제안서 작성 시 먼저 로드. 본문은 텍스트 나열 금지 — 반드시 구조화 블록으로 설계.
---

# ⚡ v6.0 핵심 원칙 (먼저 읽을 것)

> **기획 단계부터 "이 슬라이드를 N개 구조화 블록으로 디자인한다"를 먼저 결정하고
> auto_ppt.py 신규 함수를 명시적으로 호출하라. 텍스트 줄글 나열은 폐기.**

1. **목차 = 텍스트 나열 금지.** `create_toc_slide(prs, sections, current=i)` 또는
   `DeckEngine.toc(sections, current=i)` 호출 → 인간 _final 어젠다카드(좌측 accent
   블록 + 로마숫자 + 섹션명 + 활성 행 빨간 윤곽 하이라이트) 자동 재현. 섹션 디바이더는
   같은 카드에서 `current`만 다음 섹션으로 이동.
2. **본문 = 구조화 블록.** 핵심 포인트가 3~4개면 줄글 대신
   `add_structured_content_blocks(slide, items)` 또는 `DeckEngine.content_blocks(...)`
   호출 → 화면을 N등분, 각 카드 = 모서리 둥근 직사각형(브랜드색 헤더밴드 + 윤곽선 +
   그림자) 자동 배치. `items=[{"title","body"|"bullets"}, ...]`.
3. **박스 안 박스 = 컨테이너.** `add_container_box(...)` / `DeckEngine.container_slide(...)`
   → 외곽 박스(회색/블루) 안에 내부 텍스트 블록 N개 균등 배치.
4. **베이스 = 깡통 아님.** `DeckEngine`은 기아 _final을 베이스 템플릿으로 복제하여
   master/layout(표지·목차·본문)·테마·폰트를 100% 상속한 위에 덧그린다.

```python
from auto_ppt import DeckEngine
eng = DeckEngine()                          # 기아 _final 자동 상속(없으면 롯데 폴백)
eng.cover("기아 중장기 보상체계 개선 추진", "- 제안서 -", "2025. 09")
eng.toc(["제안 배경","수행 방안","일정 및 조직"], current=0)   # 어젠다카드
eng.content_blocks("Why HCG", "3대 핵심 역량",
    [{"title":"자동차 산업 이해","bullets":["현대/기아 계열사 다수","산업 보상 트렌드"]},
     {"title":"대기업 보상 경험","bullets":["대기업 프로젝트 다수","세분화 전문성"]},
     {"title":"직군 차별화 노하우","bullets":["직군 특화 방법론","Key Question 추진"]}])
eng.save()
```

---

# HCG 제안서 PPT 디자인 가이드라인 v6.0

> **분석 기반:** 롯데알미늄 _final PPTX XML 전수 추출 (v3 대비 추가 분석)
> - 25장 전수 shape 카운트 / 레이아웃 매핑 / 색상 추출
> - TOC, Pain Point, Overview 슬라이드 shape 좌표 실측
> - 슬라이드 순서/섹션 구조 완전 확인

---

## 0. 파일 스펙

| 항목 | 값 |
|------|----|
| 슬라이드 크기 | **10.833" × 7.5"** |
| 레이아웃 수 | **4개** (표지/목차/본문/End) |
| 본문 배경 | **흰색** |
| 총 슬라이드 | **25장** |

---

## 1. 실제 25장 구조 (확인 완료)

| # | 레이아웃 | 제목 | 섹션 |
|---|---------|------|------|
| S01 | 표지(0) | 롯데알미늄 직무기반 HR제도 설계 및 도입 | — |
| S02 | **목차(1)** | — | Ⅰ active |
| S03 | 본문(2) | Project Overview | 추진 내용 |
| S04 | 본문(2) | Client's Needs & Pain Point | 직무체계 / 평가보상 |
| S05 | 본문(2) | HCG's Approach | HR 전문가 지식 경험 |
| S06 | 본문(2) | [참고] | 高맥락 컨설팅 접근 방식 |
| S07 | 본문(2) | 유사 프로젝트 수행 사례 | 다수 기업 HR 컨설팅 |
| S08 | 본문(2) | 직무체계 개선 Overview | 직무체계 개선 |
| S09 | 본문(2) | 직무체계 개선 Process | 직무체계 개선 |
| S10 | 본문(2) | 직무체계 표준화 | 직무체계 개선 |
| S11 | 본문(2) | [참고] | Global Job Skill 분석 결과 |
| S12 | 본문(2) | 평가제도 개선 Overview | 평가제도 개선 |
| S13 | 본문(2) | 직군별 차별화 평가모델 | 평가제도 개선 |
| S14 | 본문(2) | [참고] | AI 기반 업무 Check-in |
| S15 | 본문(2) | 평가운영체계 | 평가제도 개선 |
| S16 | 본문(2) | 보상제도 개선 Overview | 보상제도 개선 |
| S17 | 본문(2) | 보상 지향점 / 정책선 설정 | 보상제도 개선 |
| S18 | 본문(2) | 보상제도 설계 | 보상제도 개선 |
| S19 | 본문(2) | 보상 Simulation | 보상제도 개선 |
| S20 | **목차(1)** | — | Ⅱ active |
| S21 | 본문(2) | Overview | 휴먼컨설팅그룹 |
| S22 | 본문(2) | 사업 영역 | HCG |
| S23 | 본문(2) | 컨설팅 영역 | HCG |
| S24 | 본문(2) | 주요 고객사 | HCG (870여 곳) |
| S25 | End(3) | — | — |

> **핵심:** 목차는 S02, S20 두 번 등장. 섹션은 Ⅰ(Overview) / Ⅱ(HCG 소개) 2개만.
> 전통적 섹션 구분 슬라이드 없음 — 모든 콘텐츠 슬라이드가 본문(Layout 2) 사용.

---

## 2. 테마 색상 (실측)

| 이름 | RGB | 용도 |
|------|-----|------|
| HCG_RED | **#921F0B** | 제목 텍스트 (accent5 × lumMod50%) |
| CORAL | **#F16249** | accent5, 코랄/살몬 강조 |
| CORAL_DK | **#400A07** | accent5 × lumMod20% = 우측 항목 배경 |
| CORAL_MED | **#B23B25** | accent5 × lumMod60% = 요약 항목 |
| GRAY | **#919191** | bg2 = 중간 회색, 헤더 바 |
| GRAY_DK | **#1D1D1D** | bg2 × lumMod20% = 좌측 항목 배경 |
| GRAY_LT | **#D9D9D9** | bg1 × lumMod85% = 연회색 배경 |
| BLUE | **#356CB5** | accent3 = 프로세스 단계 박스 |
| WINE | **#794039** | TOC 활성 섹션 번호 배경 |
| WHITE | **#FFFFFF** | bg1, accent1 |
| BLACK | **#000000** | dk1 = 선, 테두리 |

---

## 3. 레이아웃별 디자인

### Layout 1: 목차 (TOC)

레이아웃 자체 shape: `TextBox "CONTENTS"` at (0.944", 1.759"), 1.090"×0.323", bg2 fill

슬라이드 위에 수동 추가:

| 요소 | x | y (Ⅰ) | y (Ⅱ) | w | h | 스타일 |
|------|---|-------|-------|---|---|--------|
| 섹션 번호 원형 | 2.120" | 2.002" | 2.868" | 0.906" | 0.866" | active=#794039, inactive=white |
| 섹션 제목 텍스트 | 3.457" | 2.280" | 3.147" | 2.230" | 0.309" | 16pt, active=wine bold, inactive=black |
| 저작권 텍스트 | 2.120" | 3.829" | — | 8.020" | 0.254" | 6pt, bg1 fill |
| 수직 연결선 | 2.182" | 2.002" | 2.868" | — | — | CONN at section y |

> 섹션 제목 y = 섹션 번호 y + 0.278" (= number circle 내부 수직 중앙)
> S02 = Ⅰ active, S20 = Ⅱ active

### Layout 2: 본문

| 요소 | x | y | w | h | 스타일 |
|------|---|---|---|---|--------|
| 제목 ph (idx=0) | 0.575" | 0.250" | 9.729" | 0.305" | 12pt bold, 자동 HCG_RED |
| 서브헤드 ph (idx=10) | 0.575" | 0.548" | 9.728" | 0.361" | 15pt Semilight |
| 헤더 바 | 0.691" | 1.595" | 9.451" | 0.333" | bg2(#919191) fill |
| 수평 구분선 | 0.691" | 1.928" | 9.451" | 0" | dk1 선 |
| 좌컬럼 items | 0.691" | 1.944" | 4.528" | 0.496" | 간격 0.589" |
| 우컬럼 items | 5.615" | 1.944" | 4.528" | 0.496" | 간격 0.589" |
| 중간 연결 박스 | 4.752" | 1.994" | 1.251" | 0.394" | bg2(gray) |

---

## 4. 슬라이드 타입별 도식화 패턴

### 4-1. TOC 슬라이드 (S02, S20)

```
CONTENTS (레이아웃 제공)

[■ Ⅰ]  Project Overview     ← active: #794039, bold
 
[□ Ⅱ]  HCG 소개             ← inactive: white

저작권 텍스트 (6pt)
```

### 4-2. 2컬럼 대조 슬라이드 (Pain Point, S04)

```
[헤더 바 bg2 = 전통적 Consulting    │    HCG Approach]
─────────────────────────────────────────────────────
[항목1 #1D1D1D dark gray]  [연결]  [항목1 #400A07 dark red]
[항목2 #1D1D1D dark gray]  [연결]  [항목2 #400A07 dark red]
...
[요약 #D9D9D9 light gray]           [요약 #B23B25 med coral]
```
- 좌: GRAY_DK (#1D1D1D), 우: CORAL_DK (#400A07), 텍스트 모두 흰색
- 중간 연결박스: GRAY (#919191)

### 4-3. 프로세스 단계 슬라이드 (Overview, S03/S08/S12/S16)

```
라벨        내용 텍스트 (x=1.836", w=8.759")
(x=0.369")
           ┌─────┐ → ┌─────┐ → ┌─────┐ → ┌─────┐
           │Step1│   │Step2│   │Step3│   │Step4│  ← accent3 blue
           └─────┘   └─────┘   └─────┘   └─────┘
           [설명]      [설명]    [설명]    [설명]   ← bg2 gray
```
- 프로세스 박스: BLUE (#356CB5) fill, 1.976"×0.666", y=2.952"
- 시작 x=2.133", 간격=2.006"
- 설명 박스: GRAY fill, y=3.632", h=0.554"
- 라벨: x=0.369", w=1.250", white fill, 우정렬

### 4-4. 단일 컬럼 슬라이드 (본문 기본)

```
[헤더 바 bg2]
─────────────
라벨 │ 내용 (x=1.836", w=8.759")
라벨 │ 내용
...
```

### 4-5. 참고 슬라이드 (S06, S11, S14)

- 제목: "[참고]" prefix
- 본문: 복잡한 그리드/표/이미지 (실제는 그룹 shapes + 그림)
- 자동 생성 시: 단순화된 텍스트 박스 구현

---

## 5. 좌표 상수 (실측)

```python
# 레이아웃
TITLE_X, TITLE_Y = 0.575, 0.250    # 제목 placeholder
SUB_X, SUB_Y = 0.575, 0.548        # 서브헤드 placeholder

# 본문 그리드
HDR_X, HDR_Y, HDR_W, HDR_H = 0.691, 1.595, 9.451, 0.333
LINE_Y = 1.928
COL_L_X, COL_R_X = 0.691, 5.615
ITEM_W, ITEM_H = 4.528, 0.496
ITEM_Y0, ITEM_DY = 1.944, 0.589
MID_X, MID_W, MID_H = 4.752, 1.251, 0.394  # 중간 연결

# 단일 컬럼
LBL_X, LBL_W = 0.369, 1.250
CONTENT_X, CONTENT_W = 1.836, 8.759

# 프로세스 단계
PROC_X0, PROC_W_STEP = 2.133, 2.006
PROC_W, PROC_H, PROC_Y = 1.976, 0.666, 2.952
STEP_Y, STEP_H = 3.632, 0.554

# TOC
TOC_NUM_X = 2.120
TOC_NUM_W, TOC_NUM_H = 0.906, 0.866
TOC_TITLE_X, TOC_TITLE_W, TOC_TITLE_H = 3.457, 2.230, 0.309
TOC_SEC1_Y, TOC_SEC2_Y = 2.002, 2.868
```

---

## 6. auto_ppt.py 구현 패턴 (v4)

### 필수 색상 상수

```python
HCG_RED   = RGBColor(0x92, 0x1F, 0x0B)
CORAL_DK  = RGBColor(0x40, 0x0A, 0x07)  # Pain Point 우측
CORAL_MED = RGBColor(0xB2, 0x3B, 0x25)  # 요약 우측
GRAY      = RGBColor(0x91, 0x91, 0x91)  # 헤더/중간
GRAY_DK   = RGBColor(0x1D, 0x1D, 0x1D)  # Pain Point 좌측
GRAY_LT   = RGBColor(0xD9, 0xD9, 0xD9)  # 요약 좌측
BLUE      = RGBColor(0x35, 0x6C, 0xB5)  # 프로세스 단계
WINE      = RGBColor(0x79, 0x40, 0x39)  # TOC 활성
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
BLACK     = RGBColor(0x00, 0x00, 0x00)
```

### 핵심 헬퍼

```python
def _add_box(slide, x, y, w, h, fill_rgb=None, border_rgb=None, radius=True):
    t = MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE if radius else MSO_AUTO_SHAPE_TYPE.RECTANGLE
    s = slide.shapes.add_shape(t, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill_rgb: s.fill.solid(); s.fill.fore_color.rgb = fill_rgb
    else: _set_no_fill(s)
    if border_rgb: _set_line_color(s, border_rgb)
    else: s.line.fill.background()
    return s
```

---

## 7. 오류 레퍼런스

| 오류 | 원인 | 해결 |
|------|------|------|
| `ValueError: range 100~400000, got 0` | positional arg 순서 틀림 → fsize=0 | `_add_run(p, text, fsize, bold, ...)` 순서 확인 |
| `fill_err: cannot import PP_ALIGN from dml` | 잘못된 import | fill 함수에서 PP_ALIGN 제거 |
| 한글 폰트 깨짐 | latin/ea 미지정 | `typeface="맑은 고딕"` XML 직접 설정 |
| Git push rejected | remote ahead | `git pull --rebase origin main` 먼저 |

---

## 8. 실행

```bash
cd C:\Users\cgpar\ppt-skill
python auto_ppt.py
# → HCG_Automated_Draft.pptx 생성
```

---

## 9. 고급 도식 패턴 v4.1 (롯데알미늄 _final XML 실측)

> 신규 제안서 생성 시 평면 텍스트 대신 아래 패턴 우선. 좌표 inch. **PICTURE(아이콘)는 자동생성 제외(한계)** → 텍스트/도형 대체.

### 9-1. Pain Point 카테고리화 (S04)
| 요소 | 좌표 |
|------|------|
| 헤더 바 (좌 '고객사 HR 운영 Issue' / 우 '전통 Consulting 한계') | y=1.61 h=0.33, GRAY |
| 행 (6행) | y0=1.94 dy=0.59 h=0.50 |
| 좌 라벨태그 | x=0.69 w=1.02 h=0.50, GRAY_DK 흰 텍스트 |
| 좌/우 내용 | x=0.69 / 5.61, w=4.53 |
| 중앙 커넥터 배지 | x=4.75 w=1.25 h=0.39, GRAY |
| 하단 요약 2개 | y=5.53, w=4.53 h=1.37 |

### 9-2. VS 대비 (S05 HCG's Approach)
- 블록: 좌 x=0.69 w=4.52 / 우 x=5.65 w=4.45, y=2.20 h=4.21
- `VS` 배지: x=5.22 y=1.61 0.39², HCG_RED/WINE fill 흰 bold
- 하단 quote: x=1.10 y=6.66 w=8.64

### 9-3. 3-컬럼 STEP Overview (S08/12/16)
- 컬럼 x=[0.69, 3.90, 7.11] w=3.03
- 헤더박스 y=2.10 h=0.79 / 컨테이너 y=2.88 h=4.17 / 상단설명 y=2.94 h=0.96
- **AI 콜아웃** y=3.83 h=1.88 (예: 'AI 표준 직무체계', '직무 분류 AI Agent', '직무평가 AI Persona')
- 하단 상세 y=5.76 h=1.18 / 아이콘 0.51²(제외)

### 9-4. 직군 차별화 매트릭스 (S13)
| 열 | 좌표 |
|----|------|
| 직군 라벨 | x=0.69 w=1.25 h=0.97, HCG_RED |
| 업무 특성 | x=2.00 w=2.07 |
| 중앙 insight quote (큰따옴표) | x=4.21 w=2.68 h=0.97 |
| 화살표 배지 | x=6.99 w=0.47 h=0.57 |
| 반영방안 | x=7.57 w=2.59 |
| 행 y | 2.46 / 3.55 / 4.63 / 5.74 (사이 구분선) |
| 하단 quote | x=0.69 y=6.79 w=9.45 |

### 9-5. 예시적 배지 & Insight Quote
- `예시적` 배지: 우상단 x=9.56 y=1.61 w=0.58 h=0.27, GRAY_LT — 가설/예시 데이터 단정 회피
- Insight Quote: 하단/중앙 큰따옴표(“ ”) italic 강조색 한 줄 takeaway

### 9-6. auto_ppt.py 신규 헬퍼 (구현 지시어)
```python
add_header_bar(slide, text, x=0.69, y=1.61, w=9.451, h=0.333)   # GRAY 바 + 흰 bold
add_label_tag(slide, text, x, y, w=1.02, h=0.5, fill=GRAY_DK)   # 미니 카테고리 라벨
add_connector_badge(slide, text, x=4.752, y, w=1.25, h=0.39)    # 중앙 키워드 배지
add_vs_badge(slide, x=5.22, y=1.61, w=0.39, text='VS')          # VS 대비 배지
add_example_badge(slide, x=9.56, y=1.61, w=0.58, h=0.27)        # '예시적' 배지
add_insight_quote(slide, text, x=0.69, y=6.79, w=9.451)         # “ ” italic takeaway
build_overview_3col(prs, title, subtitle, cols)                 # 3컬럼 STEP Overview
build_diff_matrix(prs, title, subtitle, rows, bottom_quote)     # 직군 차별화 매트릭스
```

### 9-7. 한계
- 아이콘/이미지(PICTURE) 자동생성 불가 → 도형/텍스트 대체 또는 템플릿 슬라이드 복제
- 다단계 GROUP/FREEFORM 화살표 → 단순 도형 조합 근사
- 직무분류 예시 표 → add_table 또는 텍스트 그리드

---

## 10. 밀도·도식화 타깃 v4.2 (기아 _final 비교)

> 근거: **AI Draft 5.5도식/장·213자/장** vs **인간 _final 23도식/장·502자/장**. AI는 GROUP·PICTURE·LINE·TABLE·FREEFORM **0개** = 시각적 평면.

### 10-1. 타깃
| 항목 | 기준 |
|------|------|
| 도식/장 | 본문 장표 텍스트박스 나열 금지, **구조 도식 ≥1**(박스+연결선/화살표 or 표) |
| 글자/장 | 본문 **350~500자**(거버닝+근거), 단 텍스트 벽 금지 → 도식 구조화 |
| 흐름 | 텍스트 대신 **connector line/arrow** |
| 비교 | **TABLE** 활용 (인간본 8개 사용) |
| 아이콘 | PICTURE 자동생성 불가 → **add_icon_placeholder**(도형) 대체 |
| 스케일 | 진단/BM형 대형 제안서 **40~60장**, 섹션 divider 풀비주얼 |

### 10-2. auto_ppt.py 신규 헬퍼 v2 (구현 지시어)
```python
add_connector_line(slide, x1, y1, x2, y2, color=MED_GRAY, w_pt=1.0)  # 박스 간 연결선
add_arrow_flow(slide, centers)                                       # 박스 중심 사이 → 화살표 자동
add_table(slide, x, y, w, h, data, header=True)                      # 비교 TABLE (헤더 HCG_RED)
add_icon_placeholder(slide, x, y, d=0.5, fill=BLUE)                  # 아이콘 대체 원형 도형
build_process_roadmap(prs, title, subtitle, phases)                  # Process Overview 로드맵(phase+step+연결선)
build_compare_table(prs, title, subtitle, headers, rows, example=False)  # 표 기반 비교 장표
```

### 10-3. 적용 원칙
- 본문 = 미니리포트: 거버닝 메시지 + 도식화된 근거. 단계는 화살표 흐름, 관계는 연결선, 비교는 표.
- 평면 텍스트 4줄 장표 발견 시 → 도식(박스+연결선/표)으로 전환.

---

## 11. XML 레벨 고급 디자인 (v5.0)

> **근거:** 인간 _final.pptx OxmlElement 전수 파싱. AI Draft와의 갭이 평면화의 핵심 원인.

### 11-0. 실측 갭 (인간 _final vs AI Draft, 25장)
| XML 속성 | 인간 _final | AI Draft | 보강 헬퍼 |
|---|---|---|---|
| 텍스트 윤곽선 `<a:rPr><a:ln>` | **1807** | 0 | `set_text_outline` |
| 그림자 `<a:outerShdw>` | 60 | 0 | `add_shadow` |
| 투명도 `<a:alpha>` | 2478 | 0 | `set_transparency` |
| 그라디언트 `<a:gradFill>` | 43 | 0 | `set_gradient` |
| 자간 `spc` | -30~-100 | 없음 | `set_char_spacing` |
| 행간 | 100~120% 가변 | 112% 고정 | `_set_line_spacing(p,pct)` |
| 도형 geom | 14종 | 3종 | geometry_palette |
| 도형/장 · Group · Picture | 18.7 · 130 · 51 | 10.7 · 0 · 0 | 밀도/도식 강화 |

### 11-1. XML 속성 제어 규칙 + 헬퍼
```python
# 텍스트 윤곽선 — <a:rPr> 첫 자식 <a:ln w><a:solidFill><a:srgbClr></a:ln>
set_text_outline(run, color="FFFFFF", width_pt=0.75)
add_item(..., outline={"color":"FFFFFF","width_pt":1.0})   # 대비 약한 강조 텍스트만

# 그림자 — <a:spPr><a:effectLst><a:outerShdw>(ln 뒤)
add_shadow(shape, blur_pt=4, dist_pt=2.5, direction=2700000, alpha_pct=55)
add_item(..., shadow=True)                                  # 핵심 박스/배지만

# 투명도 — solidFill srgbClr에 <a:alpha val=(100-pct)*1000>
set_transparency(shape, pct)                                # 겹침·오버레이 표현

# 그라디언트 — <a:gradFill>(prstGeom 뒤·ln 앞)
set_gradient(shape, color1, color2, angle_deg=90)
add_item(..., grad=(HCG_RED, CORAL_DK, 45))                 # 표지/섹션/핵심 박스

# 자간 — <a:rPr spc='-50'>(1/100 pt, 음수=좁힘)
set_char_spacing(run, -50)                                  # 본문 -30, 제목 -50~-100
add_item(..., spc=-50)

# 세로정렬 — <a:bodyPr anchor='ctr'>
set_text_anchor(shape, "ctr")
add_item(..., anchor="ctr")
```
> 원칙: 고급 효과는 **핵심 요소 한정**. 본문 박스 남발 = 장식 과잉(톤앤매너 위배).

### 11-2. 불릿 → 도형 변환 규칙
불릿(•) 나열 **금지** → 항목 간 관계 파악 후 도식 전환:
| 관계 | 변환 |
|------|------|
| 병렬 3~5항목 | 박스 그리드(라벨태그+내용) |
| 단계/순서 | `add_arrow_flow` / `build_process_roadmap` |
| 대조(현행 vs 개선) | `build_body_2col` / `build_approach_vs` |
| 분류/매트릭스 | `build_diff_matrix` / `add_table` |
| 인과/관계 | 박스 + `add_connector_line` |
- 판단: 한 박스 불릿 3개↑ = 도식 후보.

### 11-3. MBB 모듈러 그리드
- **안전여백:** 좌우 0.69″, 상단 제목영역 ~1.6″, 하단 ~0.5″. 콘텐츠 폭 9.451″.
- **수평 분할:** 2컬럼 좌 0.691/우 5.615(거터 0.396) · 3컬럼 [0.69,3.90,7.11] w=3.03.
- **수직 리듬:** 행 피치 dy=0.589″, 박스 h=0.496″ 반복.
- **황금비:** 박스 내부여백 좌0.07/우0.04/상하0.02″, 제목:서브:콘텐츠 ≈ 1:0.7:5, 강조 박스 1.618 비율 지향.
- **정렬:** 라벨=우정렬, 본문=좌정렬, 헤더/번호/배지=중앙+`anchor=ctr`. 신규 좌표는 기존 그리드 라인에 스냅.

### 11-4. 통합 마스터 엔진 — `DeckEngine`
**원칙:** `auto_ppt.py` 하나만 호출 + 외부 JSON spec → PPT. kia/lotte는 엔진 import 콜러로 유지(중복 cover/toc/overview는 엔진 메서드로 통합).
```python
from auto_ppt import DeckEngine
eng = DeckEngine(template=None, out="out.pptx")   # None→롯데알미늄 _final 상속
eng.render(spec); eng.save()
```
```bash
python auto_ppt.py spec.json     # 데이터 구동(meta.template/out 사용)
python auto_ppt.py               # 기본 build() (롯데알미늄 25장, 하위호환)
```
**spec 스키마:** `{"meta":{template?,out?}, "slides":[{type,label?,...}]}` — type별 필드:

| type | 필드 |
|------|------|
| cover | title, subtitle?, date? |
| toc | items:[[번호,섹션,페이지],..], title? |
| overview | subtitle, background?, scope?:[[라벨,x]], sub_details?:[[x,설명]], plan?, quote? |
| section | num, title, sub? |
| body_2col | title, subtitle, header_l, header_r, left:[], right:[] |
| body_single | title, subtitle, rows:[[라벨,내용]] |
| body_process | title, subtitle, steps:[[제목,설명]], desc? |
| overview_3col | title, subtitle, cols:[{header,desc,ai,detail}×3], bar_label?, example? |
| diff_matrix | title, subtitle, rows:[{group,trait,insight,apply}], quote?, example? |
| pain_point_categorized | title, subtitle, left_hdr, right_hdr, rows:[[좌라벨,좌내용,중배지,우내용]], summary_left?, summary_right? |
| approach_vs | title, subtitle, left_title, left:[], right_title, right:[], quote? |
| process_roadmap | title, subtitle, phases:[{name,steps:[]}] |
| compare_table | title, subtitle, headers:[], rows:[[]], example? |
| appendix | title?, subtitle?, rows:[[라벨,내용]] |
| demo_advanced | title?, subtitle? (윤곽선/그림자/그라디언트/투명도/자간 시연) |
| end | — |

---

## 12. v6.0 — 기아 _final 역엔지니어링 & Text-to-Visual 엔진

> **근거:** `기아_중장기 보상체계 개선 추진_제안서_250912_v1.1_final.pptx` 58장 전수
> shape-tree 파싱(좌표 cm 실측, 그룹 재귀, fill/line/font 추출). 분석 산출물:
> `analyze_final.py` / `final_analysis.json` / `STRUCTURE_REPORT*.md`.

### 12-1. 파일 스펙 (기아 _final 실측)
| 항목 | 값 |
|------|----|
| 슬라이드 크기 | **27.52 × 19.05 cm = 10.835″ × 7.5″** |
| 레이아웃 | **표지 / 목차 / 본문 / End of document** |
| 총 슬라이드 | 58장 (목차 레이아웃 5회 = 섹션 디바이더 겸용) |
| 주 폰트 | **맑은 고딕** (Malgun Gothic), 미세주석 Malgun Gothic Semilight |

### 12-2. 브랜드 팔레트 (기아 _final 빈도순 실측)
| 이름 | RGB | 용도 |
|------|-----|------|
| KIA_RED | **#C72128** | 시그니처 레드 — 활성 섹션 accent, 윤곽 하이라이트, 헤더밴드 |
| KIA_RED_LT | **#E5838A** | 연한 레드 — 비활성 섹션 마커 |
| ACCENT_OR | **#F16249** | 강조 폰트 코랄(56회) |
| DEEP_RED | **#C00000** | 진한 강조 레드(18회) |
| CONTAINER_GR | **#E6E6E6** | 컨테이너 박스 회색 배경 |
| SKY / STEEL / NAVY | #5D8ECF / #3265A8 / #002459 | 차트·컨테이너 블루 계열 |
| PALE_BLUE | #D0E7F8 | 옅은 블루 배경 |
| BLACK / WHITE | #000000(폰트 1087회) / #FFFFFF | 본문 텍스트 / 배경 |

### 12-3. 목차 어젠다카드 (slide 1·17·52 실측, cm)
```
외곽 카드      x=5.40  y=4.93  w=20.37  h=12.02     (white, 0.5pt border, 그림자)
섹션 행(N개)   x=5.79  w=19.98  h=3.54  contiguous  (pitch 3.54, 세로 중앙정렬)
  ├ 좌 accent  x=5.79  w=2.74   h=3.54  활성=C72128 / 비활성=E5838A
  ├ 로마숫자   x=6.77  w=0.79   18pt bold 흰색 (Ⅰ Ⅱ Ⅲ …)
  └ 섹션명     x=10.54 w=12.93  18pt, 활성=C72128 bold / 비활성=black
활성 하이라이트 x=6.16 w=19.61 h=3.31  빨간(C72128) 윤곽 0.75~1.5pt, current 행 위
```
- 섹션 디바이더 = 동일 카드에서 `current`만 다음 섹션으로 이동(하이라이트 하강).
- 구현: `create_toc_slide(prs, sections, current, layout_idx, title=None)`.

### 12-4. 구조화 카드 그리드 (Text-to-Visual 계산식)
```
가용폭 W = 24.71″(=본문 마진 x=1.46cm/0.575″ 기준), n개 카드, gap=0.26″
card_w  = (W - gap*(n-1)) / n
카드     = ROUNDED_RECTANGLE, white fill, accent 1pt 윤곽, 그림자
헤더밴드  = 카드 상단 header_h=0.62″, accent(KIA_RED) 채움 + 흰 bold 타이틀(번호 옵션)
본문영역  = 헤더 아래, padding 0.13″, bullets 또는 단락, 행간 118%
n≥4 → title 11/body 9.5pt 자동축소, n≥5 → 10/9pt
```
- 구현: `add_structured_content_blocks(slide, items, x?, y=2.05, w?, h=4.30, gap=0.26, accent?, numbered=True)`.
- `items=[{"title":..,"body":..}]` 또는 `{"title":..,"bullets":[..]}` 또는 문자열.

### 12-5. 컨테이너(박스 안 박스, slide 3·31 실측)
- 외곽 SOLID 박스(E6E6E6/SKY) 안에 내부 TextBox N개 균등 분할(가로).
- 구현: `add_container_box(slide, x, y, w, h, inner_items, fill?, title?)`.
  `inner_items=[{"head":..,"body":..}, ...]`. 좌표 cm는 `CM(v)` 헬퍼로 변환.

### 12-6. DeckEngine v6.0 신규 메서드 & spec 타입
```python
eng.toc(sections, title=None, current=None)        # 어젠다카드 (튜플 입력 호환)
eng.section_divider(sections, current)             # 섹션 디바이더(하이라이트 이동)
eng.content_blocks(title, subtitle, items, ...)    # N등분 카드 그리드 본문
eng.container_slide(title, subtitle, containers)   # 박스 안 박스
```
| type | 필드 |
|------|------|
| toc | items:["섹션명",..] (또는 [[번호,섹션,페이지]] 호환), title?, current? |
| section_agenda | sections:["섹션명",..], current:int |
| blocks | title, subtitle, items:[{title,body|bullets}], y?, h?, numbered?, bar_label?, quote?, example? |
| container | title, subtitle, containers:[{x,y,w,h,fill?,title?,items:[{head,body}]}] |

> 실행: `python auto_ppt.py` → v6.0 쇼케이스(기아 템플릿, 어젠다카드+카드그리드+컨테이너)
> `python auto_ppt.py legacy` → 롯데 25장(하위호환) · `python auto_ppt.py spec.json` → 데이터 구동

---

*v6.0 — 2026-06-18 | v5.0 + 기아 _final 58장 역엔지니어링 + Text-to-Visual 엔진(목차 어젠다카드 create_toc_slide / N등분 카드 그리드 add_structured_content_blocks / 컨테이너 add_container_box) + 기아 브랜드 팔레트(C72128) + 베이스템플릿 상속 아키텍처*
*JSON 페어: skill_ppt_design.json | 기획 페어: skill_ppt_planning*
