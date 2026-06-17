---
name: hcg-ppt-design
description: >
  HCG(휴먼컨설팅그룹) 제안서 PPT 디자인 & 자동생성 가이드라인 v3.0.
  롯데알미늄 _final PPTX XML 전수 추출 + 테마 색상 직접 분석으로 정확도 극대화.
  실제 PPTX를 템플릿으로 사용하는 auto_ppt.py v3 구현 패턴 포함.
  신규 제안서 작성 또는 리뷰 시 이 스킬을 먼저 로드할 것.
---

# HCG 제안서 PPT 디자인 가이드라인 v3.0

> **분석 기반:** 롯데알미늄 _final PPTX XML 전수 추출 (테마/레이아웃/shape/폰트/색상)
> python-pptx + zipfile 직접 파싱으로 실제 값 확인
> auto_ppt.py v3 = 실제 PPTX 템플릿 기반 재생성 구현 완료

---

## 0. 파일 스펙 (절대 준수)

| 항목 | 실제 값 | 비고 |
|------|---------|------|
| **슬라이드 크기** | **10.833" × 7.5"** | US 와이드(13.33") 아님 |
| **레이아웃 수** | **4개** | 표지/목차/본문/End of document |
| **테마 파일** | theme1.xml (slide master 참조) | |
| **본문 배경** | **흰색** (master bg1 = #FFFFFF) | 배경 이미지/색상 없음 |

---

## 1. 테마 색상 실제 RGB 값

| 테마 키 | 이름 | RGB | 용도 |
|---------|------|-----|------|
| dk1 (tx1) | TEXT_1 | **#000000** | 선, 테두리, 일부 텍스트 |
| lt1 (bg1) | BG_1 | **#FFFFFF** | 슬라이드 배경, 텍스트 |
| lt2 (bg2) | BG_2 | **#919191** | 본문 body 텍스트 색상 |
| accent1 | — | **#FFFFFF** | 커버 텍스트 (outline에 사용) |
| accent2 | — | **#A1D1F1** | 연파랑 강조 |
| accent3 | — | **#356CB5** | 중파랑 강조 |
| accent4 | — | **#FFCC66** | 황금색 강조 |
| accent5 | — | **#F16249** | 코랄/살몬 강조 |
| accent6 | — | **#50B8B6** | 틸 강조 |

> **핵심:** `accent5 × lumMod=50000(50%)` ≈ **#921F0B** (다크 레드) = HCG 타이틀 텍스트 색상
> `lt2 = #919191` = 본문 body 텍스트 기본색 (회색, 흰 배경에 가독성 확보)

---

## 2. 폰트 시스템

| 항목 | 값 |
|------|-----|
| **Major font** | **맑은 고딕** (latin + EA 모두) |
| **Minor font** | **맑은 고딕** (latin + EA 모두) |
| **+mn-ea 참조** | → 맑은 고딕 (동아시아 마이너) |
| **+mj-ea 참조** | → 맑은 고딕 (동아시아 메이저) |

모든 텍스트 latin/EA 모두 `맑은 고딕` 명시. 코드에서: `typeface="맑은 고딕"`

---

## 3. 레이아웃별 디자인

### Layout 0: 표지 (`표지`)

| 요소 | 위치 | 크기 | 스타일 |
|------|------|------|--------|
| 커버 이미지 | (0.131", 0.118") | 10.571" × 2.867" | layout에서 상속 |
| Strictly Confidential | (0.127", 3.002") | 1.166" × 0.236" | Arial 8pt italic, #FFFFFF 65% lum |
| 블랙 템플릿 바 | (1.135", 4.042") | 8.564" × 0.474" | fill: tx1(black), layout |
| 회사/저작권 바 | (0.828", 6.976") | 9.177" × 0.269" | fill: tx1(black), layout |
| **타이틀 placeholder** | **(1.135", 3.559")** | 8.564" × 0.474" | 22pt bold 맑은 고딕, 검정 |
| **서브타이틀 placeholder** | **(1.135", 4.523")** | 8.564" × 0.305" | 검정 |
| **날짜 텍스트** | **(1.137", 5.319")** | 8.563" × 0.3" | 10.5pt 맑은 고딕, CENTER |

> 표지 배경 = 흰색. 커버 이미지가 상단 2.867" 채움. 하단은 흰 배경.
> 타이틀 텍스트는 이미지 아래 흰 영역에 검정으로 배치.

### Layout 2: 본문 (`본문`)

| 요소 | 위치 | 크기 | 스타일 |
|------|------|------|--------|
| **제목 placeholder** | **(0.575", 0.250")** | 9.729" × 0.305" | 12pt bold 맑은 고딕, **color≈#921F0B** (accent5×50% lum) |
| **서브헤드 placeholder (idx=10)** | **(0.575", 0.548")** | 9.728" × 0.361" | 15pt 맑은 고딕 Semilight, 텍스트 outline 1pt, 거의 검정 |

> 본문 레이아웃 background = 없음 (흰색 상속). 컬러 헤더 바 없음.
> 제목은 다크 레드 텍스트 (배경 아님). 서브헤드는 텍스트 윤곽선 포함.

---

## 4. 본문 콘텐츠 배치 좌표 (실측)

| 요소 | x (inch) | y 시작 | width | height | 간격(y) |
|------|---------|--------|-------|--------|--------|
| 좌측 컬럼 아이템 | **0.691** | **1.944** | **4.528** | **0.496** | **0.589** |
| 우측 컬럼 아이템 | **5.615** | **1.944** | **4.528** | **0.496** | **0.589** |
| 컬럼 헤더(색상) | 0.691 / 5.615 | **1.606** | 4.528 | **0.333** | — |
| 중간 연결(VS 등) | **5.22** | row_y+0.05 | **0.394** | **0.394** | — |
| 행 라벨 (우측정렬) | **0.369** | row_y | **1.250** | 0.318 | — |
| 넓은 콘텐츠 박스 | **1.836** | row_y | **8.759** | 0.496 | — |

### 아이템 y 위치 (6행 기준)

| 행 | y |
|----|---|
| 1 | 1.944 |
| 2 | 2.533 |
| 3 | 3.122 |
| 4 | 3.712 |
| 5 | 4.301 |
| 6 | 4.891 |

---

## 5. Shape 스타일

### 본문 콘텐츠 아이템 (Rounded Rectangle)

```python
shape = slide.shapes.add_shape(
    MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
    Inches(0.691), Inches(1.944), Inches(4.528), Inches(0.496)
)
# fill: noFill (transparent)
# line: 0.5pt, dk1(#000000) color
# font: 맑은 고딕 11pt, color #919191 (lt2)
# margin: left 0.07", right 0.04", top/bot 0.02"
# line_spacing: 112%
```

### 컬럼 헤더 (색상 배경)

```python
# fill: HCG_RED = #921F0B or DARK_GRAY = #404040
# line: noFill
# font: 맑은 고딕 11pt bold WHITE
# height: 0.333" (아이템보다 낮음)
```

### 라벨 텍스트 (행 구분)

```python
# font: 맑은 고딕 13pt italic
# color: MED_GRAY = #919191
# alignment: RIGHT
# no fill, no border
```

---

## 6. auto_ppt.py 구현 패턴 (v3 핵심)

### 템플릿 사용법

```python
# 실제 PPTX를 템플릿으로 로드 — 테마/폰트/레이아웃 완전 상속
prs = Presentation("..._final.pptx")

# 기존 슬라이드 전부 삭제
sldIdLst = prs.slides._sldIdLst
for sld_el in list(sldIdLst):
    rId = sld_el.get(qn('r:id'))
    prs.part.drop_rel(rId)
    sldIdLst.remove(sld_el)

# 새 슬라이드 추가 (레이아웃 인덱스)
# 0: 표지, 1: 목차, 2: 본문, 3: End of document
slide = prs.slides.add_slide(prs.slide_layouts[2])
```

### Placeholder 텍스트 설정

```python
for ph in slide.placeholders:
    idx = ph.placeholder_format.idx
    if idx == 0:           # title → 자동 dark red 12pt bold
        ph.text = "슬라이드 제목"
    elif idx == 10:        # body subtitle → 자동 15pt Semilight
        ph.text = "서브헤드 텍스트"
```

### Korean 폰트 명시 설정

```python
def _set_font_xml(run, name="맑은 고딕"):
    rPr = run._r.get_or_add_rPr()
    for tag in [qn('a:latin'), qn('a:ea')]:
        el = rPr.find(tag)
        if el is None:
            el = etree.SubElement(rPr, tag)
        el.set('typeface', name)
```

### noFill 설정

```python
def _set_no_fill(shape):
    spPr = shape._element.find(qn('p:spPr'))
    ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    for tag in ['solidFill','gradFill','blipFill','pattFill','grpFill','noFill']:
        for el in spPr.findall(f'{{{ns}}}{tag}'):
            spPr.remove(el)
    etree.SubElement(spPr, f'{{{ns}}}noFill')
```

### 선 색상 (dk1) 설정

```python
def _set_line_dark(shape, pt=0.5):
    spPr = shape._element.find(qn('p:spPr'))
    ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    ln_el = etree.SubElement(spPr, f'{{{ns}}}ln')
    ln_el.set('w', str(int(pt * 12700)))
    sF = etree.SubElement(ln_el, f'{{{ns}}}solidFill')
    sc = etree.SubElement(sF, f'{{{ns}}}schemeClr')
    sc.set('val', 'dk1')
```

---

## 7. 자주 쓰는 슬라이드 패턴

### 2-Column 비교 슬라이드

```
[헤더 좌: HCG_RED fill]        [헤더 우: DARK fill]
 y=1.606" h=0.333"              y=1.606" h=0.333"

[아이템 1-1]  x=0.691"          [아이템 1-2]  x=5.615"
[아이템 2-1]  y=2.533"          [아이템 2-2]
[아이템 3-1]  y=3.122"          [아이템 3-2]
...  (최대 5행)
```

### Process 슬라이드 (→ 화살표)

```
[단계1] → [단계2] → [단계3] → [단계4]
   설명        설명        설명        설명
```

### Single Column 슬라이드 (라벨+내용)

```
라벨 ▶│ 내용 텍스트 (x=1.836", w=8.759")
(우정렬│italic, 13pt, #919191)
```

---

## 8. 검증 항목 체크리스트

- [ ] 슬라이드 크기 10.833" × 7.5"
- [ ] 실제 PPTX 템플릿 기반 (`Presentation(real_file)`)
- [ ] 슬라이드 수 25장
- [ ] 제목 placeholder idx=0 사용 → 자동 dark red 스타일
- [ ] 서브헤드 placeholder idx=10 사용 (없으면 textbox 대체)
- [ ] 본문 아이템 = rounded rectangle, noFill, 0.5pt border
- [ ] 폰트 = `맑은 고딕` (latin+EA 모두 명시)
- [ ] 좌컬럼 x=0.691", 우컬럼 x=5.615"
- [ ] 아이템 높이 0.496", 폭 4.528"
- [ ] y 시작 1.944", 간격 0.589"

---

## 9. 주요 오류 및 해결

| 오류 | 원인 | 해결 |
|------|------|------|
| `ValueError: must be in range 100~400000, got 0` | positional arg 순서 오류 → fsize=0 전달 | `_add_run(para, text, fsize, bold, italic, color)` 순서 일치 확인 |
| `AttributeError: list has no attribute 'rId'` | `prs.slides[:n]` 슬라이싱 | `list(prs.slides)[:n]` 사용 |
| 한글 경로 깨짐 (PowerShell) | inline string 인코딩 | `.py` 파일로 저장 후 실행 |
| `fill_err: cannot import PP_ALIGN from dml` | 잘못된 import 위치 | fill 함수에서 PP_ALIGN import 제거 |
| Git push rejected | remote 커밋 존재 | `git pull --rebase origin main` 먼저 |

---

## 10. 실행 방법

```bash
cd C:\Users\cgpar\ppt-skill
python auto_ppt.py
# → HCG_Automated_Draft.pptx 생성
```

---

*v3.0 업데이트: 2026-06-17*
*XML 전수 추출 기반 정확도 극대화 — 실제 PPTX 템플릿 사용 방식으로 근본적 디자인 복제 구현*
