---
name: hcg-ppt-design
description: >
  HCG 제안서 PPT 디자인 & 자동생성 가이드라인 v4.0.
  롯데알미늄 _final PPTX 전수 XML 분석 — 실제 25장 구조/레이아웃/색상/좌표 완전 정합.
  도식화·구성·구조화 패턴 포함. 신규 제안서 작성 시 먼저 로드.
---

# HCG 제안서 PPT 디자인 가이드라인 v4.0

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

*v4.2 — 2026-06-18 | v4.1 + 밀도/도식화 타깃(기아 _final 23도식/장 비교) + connector/table/roadmap 헬퍼 지시어*
*JSON 페어: skill_ppt_design.json | 기획 페어: skill_ppt_planning*
