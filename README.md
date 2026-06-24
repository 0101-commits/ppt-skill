# ppt-skill — HCG 제안서 슬라이드 자동화 프레임워크

휴먼컨설팅그룹(HCG) 컨설팅 제안서를 **기획(Planning)** 과 **디자인(Design)** 두 스킬로 관리하고,
데이터 구동 프레임워크(`main.py` + `core/`)로 실제 슬라이드(**PPTX + HTML**)를 자동 생성하는 저장소.

신규 제안서는 Python 코드 없이 `projects/<client>.json` 한 개만 작성하면 된다.

> ## 디자인 시스템 v2.0 리셋 (절대 기준: `HCG-slide-design-system.md`)
> v2.0은 **HCG-Slide-Design-System v1.0** 을 절대 기준으로 한 전면 재설계다. 기존 하우스 스타일
> (4:3 / 맑은 고딕 / HCG_RED 중심)을 폐기하고 다음으로 리셋했다.
>
> | 항목 | v1.5 (legacy) | **v2.0 (현재)** |
> |------|---------------|-----------------|
> | 캔버스 | 4:3 · 10.833″×7.5″ | **16:9 · 1280×720 px** (px 절대 좌표, `PX()=Inches(px/96)`) |
> | 폰트 | 맑은 고딕 (latin/ea 혼용) | **Pretendard 단일** (혼용 금지) |
> | 메인 색 | HCG_RED `#921F0B` | **메인 블루 `#356CB5`** (화면 60–70%). 마룬은 `HCG` 워드마크 전용 |
> | 골격 | 레이아웃 4종(표지/목차/본문/End) | **고정 7요소 골격**(아래) — 슬라이드 불문 동일 좌표·폰트 |
> | 제목 | 명사구 라벨 | **결론형(So-What) 완결 문장** |
> | 데이터 | 표/도형 중심 | **차트 우선**(수치 슬라이드 ≥70% 차트), **출처 상시 표기** |
> | 밀도 | — | **콘텐츠 점유율 ≥85%** 강제 |
> | 엔진 | PPTX(real-pptx 템플릿 상속) | **PPTX(빈 16:9 캔버스 생성) + HTML/CSS** 2종 |
>
> **고정 7요소 골격:** `kicker` → `chapter-indicator` → `slide-title` → `title-rule` →
> `content-area` → `source-note` → `page-number`. 엔진이 모든 콘텐츠 슬라이드에 자동 상속한다.
> 색 띠·밑줄 컬러바·슬라이드폭 헤더/푸터 바 금지 — 구분은 면 틴트·1px 룰·아이콘으로.

---

## 폴더 및 파일 구조

```
ppt-skill/
├── HCG-slide-design-system.md    # ★ 디자인 절대 기준 (HCG-Slide-Design-System v1.0)
├── main.py                       # 통합 CLI 진입점 (--client/--out/--both|--pptx|--html)
├── core/                         # 핵심 공통 로직 (Core)
│   ├── __init__.py
│   ├── tokens.py                 # 디자인 토큰 단일 소스(PALETTE/TYPO/layout, 순수·렌더 비의존) ← v2.0 신규
│   ├── schema.py                 # 슬라이드 타입 레지스트리(13종) + 검증 (단일 진실 공급원)
│   ├── planner.py                # 기획 엔진: config → 정규화 spec (렌더 비의존, dict→dict)
│   ├── designer.py               # 디자인 엔진(PPTX): spec → .pptx (python-pptx, 빈 16:9 캔버스)
│   └── html_renderer.py          # 디자인 엔진(HTML): spec → .html (CSS, Chart.js)  ← v2.0 신규
├── projects/                     # 프로젝트별 데이터 (Projects)
│   ├── lotte_chemical.json       # 클라이언트 아이덴티티 + 슬라이드 콘텐츠 (샘플)
│   └── hyundai.json
├── tests/                        # pytest (schema/planner/cli/integration/designer)
│   ├── test_schema.py
│   ├── test_planner.py
│   ├── test_cli.py
│   ├── test_integration.py
│   └── test_designer_smoke.py
├── docs/superpowers/             # 설계 문서 + 구현 계획
├── history/                      # 과거 이력 및 초안 (History) — 동결(frozen)
│   ├── designer_legacy_4x3.py    # ★ v1.5 4:3 엔진 동결 스냅샷 (real-pptx 템플릿 상속)
│   ├── auto_ppt_legacy.py        # 리팩터링 전 모놀리식 auto_ppt.py 스냅샷
│   ├── auto_ppt_kia.py / lotte_chemical_ppt.py / paradise_compare.py  # 명령형 덱 생성기
│   ├── analyze_final.py / report2.py / report_final.py   # 일회성 분석 스크립트
│   ├── *.json (final_analysis, spec_paradise, paradise_compare)
│   └── STRUCTURE_REPORT*.md / PARADISE_COMPARE.md
├── skill_ppt_planning.json       # 기획 스킬 — JSON (planner가 읽음)        [v2.0]
├── skill_ppt_planning.md         # 기획 스킬 — MD (사람 판독 정본)          [v2.0]
├── skill_ppt_design.json         # 디자인 스킬 — JSON (디자인 contract 정본) [v2.0]
├── skill_ppt_design.md           # 디자인 스킬 — MD (사람 판독 정본)        [v2.0]
└── README.md                     # (이 파일)
```

세 계층 분류:
- **Core** (`core/`, `main.py`) — 어떤 프로젝트든 공유하는 엔진/CLI.
- **Projects** (`projects/*.json`) — 클라이언트별 데이터. 코드가 아닌 선언적 설정.
- **History** (`history/`) — 리팩터링 전 레거시 스크립트/분석 산출물. 참고·복원용으로 동결.

> `HCG-slide-design-system.md` 가 디자인의 **절대 기준**이고, `skill_ppt_design.json` 은 그 frontmatter
> 토큰을 기계 판독용으로 인코딩한 **디자인 contract 정본**(slide_types / engine / cli)이다. 두 파일은 항상 동기.

---

## 사용법 (통합 프레임워크)

```bash
cd C:\Users\cgpar\ppt-skill

python main.py --client lotte_chemical               # 기본 --both → .pptx + .html 동시 생성
python main.py --client lotte_chemical --both        # PPTX + HTML 모두
python main.py --client lotte_chemical --pptx        # PPTX만
python main.py --client lotte_chemical --html        # HTML만
python main.py --client lotte_chemical --out deck     # 출력 경로 override (확장자 자동)
```

- CLI 서명: `python main.py --client <name> [--out path] [--both|--pptx|--html]`.
- 출력 포맷 미지정 시 **기본 `--both`** — 동일 spec에서 `.pptx`(`core/designer.py`)와
  `.html`(`core/html_renderer.py`)을 함께 emit 한다.

**의존성:** Python 3.10+, `python-pptx`

```bash
pip install python-pptx pytest
```

- **Pretendard** — PPTX 렌더링을 위해 시스템 설치 또는 폰트 임베드 필요(미설치 시 PowerPoint 폴백 폰트로
  표시). HTML은 Pretendard `@font-face` 를 CDN(`jsdelivr`)에서 로드 — 열람 시 네트워크 필요(오프라인은 자가 호스팅 권장).
- **Chart.js** — HTML 슬라이드 차트는 Chart.js를 CDN에서 로드(폰트 Pretendard). PPTX는 python-pptx 네이티브 차트 사용.

---

## 새 클라이언트 추가 (Python 코드 불필요)

`projects/<client>.json` 한 개만 작성한다.

```json
{
  "client": "lotte_chemical",
  "identity": {
    "display_name": "롯데케미칼"
  },
  "content": {
    "out": "HCG_롯데케미칼_제안서_Draft",
    "slides": [
      {"type": "cover", "title": "직무·성과 기반 보상체계 재설계 제안",
       "subtitle": "- 제안서 -", "date": "2025.06"},
      {"type": "toc", "items": ["제안 개요", "현황 진단", "추진 방안", "실행 로드맵"], "current": 0},
      {"type": "section", "title": "현황 진단"},
      {"type": "compare", "title": "연공 중심 보상구조가 성과 변별력을 저해함",
       "left_header": "As-Is", "left": ["..."], "right_header": "To-Be", "right": ["..."]},
      {"type": "chart", "title": "직급별 인건비가 5년간 32% 증가함",
       "chart": {"type": "bar", "labels": ["2021","2022","2023","2024","2025"],
                  "series": [{"name": "인건비", "data": [100,108,119,127,132]}], "unit": "지수"},
       "source": "전자공시시스템 사업보고서, 2024"},
      {"type": "kpi", "title": "핵심 성과 지표",
       "items": [{"number": "32", "unit": "%", "label": "인건비 증가율"}]},
      {"type": "end"}
    ]
  }
}
```

- `identity` → `spec.meta`, `content.slides` → `spec.slides`.
- `identity`는 **`display_name`** 을 가진다. v1.5의 per-client `theme`/`colors`(HCG_RED)/`fonts`(맑은 고딕)
  지정은 **제거**됐다 — 팔레트/폰트는 `skill_ppt_design.json` 의 디자인 토큰에서 단일하게 적용된다.
- **`template`은 더 이상 필요 없다**(엔진이 빈 16:9 캔버스에서 생성). 남아 있어도 **선택/무시**.
- 모든 슬라이드는 명시적 `type`을 가져야 한다(선언적; 추론 없음). 알 수 없는 타입은 즉시 에러.
- 콘텐츠 타입은 `kicker?` / `chapter?` / `source?` 를 받아 고정 골격에 채운다. `title`은 **결론 문장(필수)**.

### 슬라이드 타입 (v2.0 — `skill_ppt_design.json` `slide_types`)

| type | 필수 키 | 용도 |
|------|---------|------|
| `cover` | `title` | 표지: hero-band + 중앙 정렬 title/subtitle/date + HCG 워드마크 + 하단 저작권 |
| `toc` | `items` | 목차/섹션 디바이더(로마숫자 칩). `current`=강조 인덱스 |
| `section` | `title` | 섹션 간지(풀비주얼). `num` 자동 로마숫자 |
| `bullets` | `title`, `items` | ■ 불릿 본문(2단까지). `items=[str]` 또는 `[{text, sub:[..]}]` |
| `cards` | `title`, `items` | content-card 그리드(2/3/4-up). `items=[{header, body|bullets}]` |
| `columns` | `title`, `columns` | 칼럼 그리드 + column-header(/-dark). `columns=[{header, dark?, body|bullets}]` |
| `compare` | `title`, `left`, `right` | 2-up As-Is/To-Be. To-Be 헤더 column-header-dark |
| `kpi` | `title`, `items` | kpi-stat 행(4-up). `items=[{number, unit?, label}]` |
| `process` | `title`, `steps` | process-flow 노드 + 화살표. `steps=[{label, desc?}]` |
| `matrix` | `title`, `columns`, `rows` | 차별화 매트릭스. `rows=[{group, cells:[..]}]` |
| `chart` | `title`, `chart` | 차트 슬라이드. `chart={type, labels:[..], series:[{name, data:[..]}], unit?}` |
| `table` | `title`, `headers`, `rows` | data-table. `emphasis_row`=강조 행 인덱스 |
| `end` | — | 마무리 슬라이드. `message?` |

> 각 타입의 선택 키 전체는 `skill_ppt_design.json` 의 `slide_types`, 검증 규칙은 `core/schema.py` 참고.

---

## 데이터 흐름

```
projects/<client>.json
   ├─ identity{display_name}
   └─ content{out, slides[]}
                 │
      Planner.plan()   ← skill_ppt_planning.json (스토리라인/섹션 규칙)
                 │  (검증 · section 로마숫자 자동번호 · meta 구성)
                 ▼
      정규화 spec {meta, slides[]}
                 │
                 ├──────────────────────────────┐
   Designer.render(spec)            HtmlRenderer.render(spec)
   ← skill_ppt_design.json 토큰     ← skill_ppt_design.json 토큰
   (core/designer.py)               (core/html_renderer.py)
                 ▼                              ▼
            출력 .pptx                      출력 .html
```

두 엔진은 `core/tokens.py`(순수 토큰 단일 소스: `colors`/`typography`/`layout`/`charts`)를 공유하고,
실행 시 `skill_ppt_design.json` 의 `colors` 토큰을 `apply_design_tokens(json)` 으로 그 위에 주입한다.

---

## 핵심 구조: 스킬 = 2개 × 2포맷

스킬은 **기획/디자인 2개**로 나뉘고, 각각 **JSON과 MD 두 버전을 별도 관리**한다. (모두 **v2.0**)

| 스킬 | JSON (프롬프트/기계 판독) | MD (사람 판독/리뷰) | 담당 영역 |
|------|--------------------------|---------------------|-----------|
| **기획 (Planning)** | `skill_ppt_planning.json` | `skill_ppt_planning.md` | 스토리라인, 목차, 헤드 메시지, 논리 흐름, 맥락 요약 |
| **디자인 (Design)** | `skill_ppt_design.json` | `skill_ppt_design.md` | 16:9 골격, 색상/폰트 토큰, 좌표, 컴포넌트, 차트, 슬라이드 타입 |

> 한쪽을 수정하면 다른 쪽도 함께 갱신해 두 버전을 동기 상태로 유지한다.
> - **JSON** — LLM 프롬프트 주입·자동화 파이프라인용. 키 기반 구조.
> - **MD** — 사람이 읽고 리뷰·수정하는 정본. 표/코드블록/체크리스트.

### 1. 기획 (`skill_ppt_planning`)
- **표준 스토리라인:** 표지 → 목차 → `Ⅰ 프로젝트 배경` → `Ⅱ 추진방안` → `Ⅲ 일정·조직` → `Ⅳ HCG 소개` → End
- **핵심 메시지 원칙:** 제도(MBO/OKR/직무급)는 껍질, 본질은 **평가-보상 연계 구조**. 한 문장 주장으로 수렴.
- **배경 전개:** `변화 → 현상 → 가설 이슈 → HCG 관점 → 대응`(문제 나열식 금지).
- **결론형 제목:** 장표 제목은 그 슬라이드의 So-What 완결 문장(제목만 이어 읽어도 스토리 성립).

### 2. 디자인 (`skill_ppt_design` — `HCG-slide-design-system.md` 기준)
- **캔버스:** 16:9 · 1280×720 px, 흰 배경 고정. 좌표 px 절대(`PX()=Inches(px/96)`).
- **폰트:** **Pretendard 단일**(본문/제목/숫자/영문 라틴 전부). weight 400 본문 / 700 강조 2단 대비.
- **메인 색:** 메인 블루 `#356CB5`(화면 60–70%) + 라이트블루 틴트 면 + 코랄 `#F16249` 단일 액센트.
  마룬 `#921F0B`은 `HCG` 워드마크 전용. emphasis-red/positive-green은 증감 수치 전용.
- **고정 골격:** kicker / chapter-indicator / slide-title / title-rule / content-area / source-note / page-number.
- **차트 우선:** 정량 데이터는 표가 아닌 차트가 1순위. 6색 시리즈 팔레트 고정.
- **고밀도:** content-area(1168×512)를 2/3/4-up·2×2·L-rail 그리드로 분할 충전(점유율 ≥85%, gutter 24·block-gap 12).

---

## 테스트

```bash
cd C:\Users\cgpar\ppt-skill
python -m pytest -v
```

- `test_schema` — 슬라이드 타입 레지스트리, 알 수 없는 타입/누락 필드 검출.
- `test_planner` — config→spec 구조화, 섹션 로마숫자, 입력 불변, 렌더 비의존.
- `test_cli` — `--client/--out/--both|--pptx|--html` 및 친절한 에러.
- `test_integration` — 샘플 config 엔드투엔드(.pptx/.html).
- `test_designer_smoke` — 엔진 API 보존 + 디자인 토큰 소비.

---

## 분석 기반 / 연혁

- **v1.0** — 10개 제안서 530+ 파일 분석으로 통합 스킬(기획+디자인) 등록.
- **v1.x** — 디자인/좌표 정밀화(롯데알미늄 `_final` XML 전수 추출), 기아·파라다이스 `_final` 팔레트 역추출.
- **스킬 분리** — 통합 스킬을 **기획/디자인 2개 스킬 × JSON/MD 2포맷**으로 분리.
- **통합 프레임워크 (2026-06)** — 프로젝트별 명령형 스크립트(kia/lotte/paradise)를 데이터 구동
  (`projects/*.json` + `core/` + `main.py`)으로 리팩터링. 레거시는 `history/`로 동결.
- **디자인 시스템 v2.0 리셋** — **HCG-Slide-Design-System v1.0**(16:9/Pretendard/블루/고정 골격)을
  절대 기준으로 전면 재설계, **HTML 렌더러 추가**. v1.5 4:3 엔진은 `history/designer_legacy_4x3.py` 로 동결.

> 디자인 절대 기준: `HCG-slide-design-system.md`. 콘텐츠 우선 규칙: `fb_`(리더 피드백) · `_final`(최종본)
> 기준이 초안과 충돌 시 항상 `fb_`/`_final` 우선.
