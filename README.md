# ppt-skill — HCG 제안서 PPT 자동화 프레임워크

휴먼컨설팅그룹(HCG) 컨설팅 제안서를 **기획(Planning)** 과 **디자인(Design)** 두 스킬로 관리하고,
데이터 구동 프레임워크(`main.py` + `core/`)로 실제 PPTX를 자동 생성하는 저장소.

10개 제안서 약 530개 파일(피드백 `fb_`, 최종본 `_final`, 내부회의록 포함)을 교차 분석해 도출한
HCG 제안서 표준을 코드화했다.

> **2026-06-19 통합 프레임워크 리팩터링:** 프로젝트별 명령형 스크립트(kia/lotte/paradise)를
> **데이터 구동 구조**로 전환. 신규 제안서는 Python 코드 없이 `config/<client>.json` 한 개만 작성하면 된다.
> 검증된 렌더링 엔진(`auto_ppt.py`)은 `core/designer.py`로 **그대로 보존**했다.

---

## 폴더 및 파일 구조

```
ppt-skill/
├── main.py                       # 통합 CLI 진입점 (--client/--list/--dry-run/--validate/--out)
├── core/                         # 핵심 공통 로직 (Core)
│   ├── __init__.py
│   ├── schema.py                 # 19종 슬라이드 타입 레지스트리 + 검증 (단일 진실 공급원)
│   ├── planner.py                # 기획 엔진: config → 정규화 spec (pptx 비의존, 순수 dict→dict)
│   └── designer.py               # 디자인 엔진: spec → pptx (검증된 DeckEngine, 실측 좌표/테마)
├── config/                       # 프로젝트별 데이터 (Projects)
│   └── lotte_chemical.json       # 클라이언트 아이덴티티 + 슬라이드 콘텐츠 (샘플)
├── tests/                        # pytest (schema/planner/cli/integration/designer)
│   ├── test_schema.py
│   ├── test_planner.py
│   ├── test_cli.py
│   ├── test_integration.py
│   └── test_designer_smoke.py
├── docs/superpowers/             # 설계 문서 + 구현 계획
│   ├── specs/2026-06-19-ppt-skill-refactor-design.md
│   └── plans/2026-06-19-ppt-skill-unified-framework.md
├── archive_history/              # 과거 이력 및 초안 (History) — 동결(frozen)
│   ├── auto_ppt_legacy.py        # 리팩터링 전 모놀리식 auto_ppt.py 스냅샷
│   ├── auto_ppt_kia.py           # 기아 명령형 덱 생성기
│   ├── lotte_chemical_ppt.py     # 롯데케미칼 명령형 덱 생성기
│   ├── paradise_compare.py       # 파라다이스 Draft-vs-final QA 도구
│   ├── analyze_final.py / report2.py / report_final.py   # 일회성 분석 스크립트
│   ├── *.json (final_analysis, spec_paradise, paradise_compare)
│   └── STRUCTURE_REPORT*.md / PARADISE_COMPARE.md
├── skill_ppt_planning.json       # 기획 스킬 — JSON (planner가 읽음)
├── skill_ppt_planning.md         # 기획 스킬 — MD (사람 판독 정본)
├── skill_ppt_design.json         # 디자인 스킬 — JSON (designer가 색상/폰트 토큰 소비)
├── skill_ppt_design.md           # 디자인 스킬 — MD (사람 판독 정본)
└── README.md                     # (이 파일)
```

세 계층 분류:
- **Core** (`core/`, `main.py`) — 어떤 프로젝트든 공유하는 엔진/CLI.
- **Projects** (`config/*.json`) — 클라이언트별 데이터. 코드가 아닌 선언적 설정.
- **History** (`archive_history/`) — 리팩터링 전 레거시 스크립트/분석 산출물. 참고·복원용으로 동결.

---

## 사용법 (통합 프레임워크)

```bash
cd C:\Users\cgpar\ppt-skill

python main.py --list                                # 등록된 클라이언트 목록
python main.py --client lotte_chemical               # 렌더링 → config의 out 경로로 PPTX 생성
python main.py --client lotte_chemical --dry-run     # 정규화 spec(JSON)만 출력, 렌더 안 함
python main.py --client lotte_chemical --validate    # config/spec 검증만, 렌더 안 함
python main.py --client lotte_chemical --out deck.pptx   # 출력 경로 override
```

### 분야별(프로젝트별) 실행

각 제안서는 `config/<client>.json` 하나로 정의되며, 같은 `main.py`로 분야별 실행한다.

```bash
# 롯데케미칼 직무기반 HR 제안서 (샘플 제공)
python main.py --client lotte_chemical

# 신규/복원 프로젝트 예시 (config/<client>.json 작성 후 동일 패턴)
python main.py --client kia                # config/kia.json
python main.py --client paradise           # config/paradise.json
python main.py --client lotte_aluminium    # config/lotte_aluminium.json
```

> 기아·파라다이스·롯데알미늄 등 과거 명령형 스크립트는 현재 `archive_history/`에 동결되어 있다.
> 복원하려면 해당 덱 콘텐츠를 `config/<client>.json`으로 옮겨 작성한 뒤 위 명령으로 렌더링한다.

**의존성:** Python 3.10+, `python-pptx`

```bash
pip install python-pptx pytest
```

---

## 새 클라이언트 추가 (Python 코드 불필요)

`config/<client>.json` 한 개만 작성한다.

```json
{
  "client": "lotte_chemical",
  "identity": {
    "display_name": "롯데케미칼",
    "theme": "hcg",
    "template": "C:\\path\\to\\base_final.pptx",
    "colors": {"HCG_RED": "#921F0B"},
    "fonts": {"kr": "맑은 고딕", "en": "Arial"}
  },
  "content": {
    "out": "HCG_롯데케미칼_제안서_Draft.pptx",
    "slides": [
      {"type": "cover", "title": "...", "subtitle": "- 제안서 -", "date": "2025.06"},
      {"type": "toc", "items": ["프로젝트 배경", "추진 방안", "..."]},
      {"type": "approach_vs", "title": "...", "subtitle": "...",
       "left_title": "...", "left": ["..."], "right_title": "...", "right": ["..."]},
      {"type": "end"}
    ]
  }
}
```

- `identity` → `spec.meta` (테마/템플릿/색상/폰트), `content.slides` → `spec.slides`.
- 모든 슬라이드는 명시적 `type`을 가져야 한다(선언적; 추론 없음). 알 수 없는 타입은 즉시 에러.
- 사용 가능한 19종 슬라이드 타입과 각 필수/선택 키는 `core/schema.py`의 `SLIDE_TYPES` 참고.
- 스키마 상세는 `docs/superpowers/specs/2026-06-19-ppt-skill-refactor-design.md` §7.

---

## 데이터 흐름

```
config/<client>.json
   ├─ identity{theme, template, colors, fonts}
   └─ content{out, slides[]}
                 │
      Planner.plan()   ← skill_ppt_planning.json (스토리라인/섹션 규칙)
                 │  (검증 · section 로마숫자 자동번호 · meta 구성)
                 ▼
      정규화 spec {meta, slides[]}
                 │
      Designer.render(spec)   ← skill_ppt_design.json (색상/폰트 토큰) + identity 색상
                 ▼
             출력 .pptx
```

---

## 핵심 구조: 스킬 = 2개 × 2포맷

스킬은 **기획/디자인 2개**로 나뉘고, 각각 **JSON과 MD 두 버전을 별도 관리**한다.

| 스킬 | JSON (프롬프트/기계 판독) | MD (사람 판독/리뷰) | 담당 영역 |
|------|--------------------------|---------------------|-----------|
| **기획 (Planning)** | `skill_ppt_planning.json` | `skill_ppt_planning.md` | 스토리라인, 목차, 헤드 메시지, 논리 흐름, 맥락 요약 |
| **디자인 (Design)** | `skill_ppt_design.json` | `skill_ppt_design.md` | 톤앤매너, 색상/폰트, 레이아웃, 좌표 상수, 도식화 패턴 |

> 한쪽을 수정하면 다른 쪽도 함께 갱신해 두 버전을 동기 상태로 유지한다.
> - **JSON** — LLM 프롬프트 주입·자동화 파이프라인용. 키 기반 구조.
> - **MD** — 사람이 읽고 리뷰·수정하는 정본. 표/코드블록/체크리스트.

### 1. 기획 (`skill_ppt_planning`)
- **표준 스토리라인:** 표지 → 목차 → `Ⅰ 프로젝트 배경` → `Ⅱ 추진방안(3모듈)` → `Ⅲ 일정·조직` → `Ⅳ HCG 소개` → End
- **핵심 메시지 원칙:** 제도(MBO/OKR/직무급)는 껍질, 본질은 **평가-보상 연계 구조**. 한 문장 주장으로 수렴.
- **배경 전개:** `변화 → 현상 → 가설 이슈 → HCG 관점 → 대응` (문제 나열식 금지).
- **표준 분량:** 17~24장(고객 요청 우선, 대형 RFP는 80~100장 허용).

### 2. 디자인 (`skill_ppt_design`)
- **파일 스펙:** 10.833" × 7.5", 레이아웃 4종(표지/목차/본문/End), 흰 배경.
- **테마 색:** HCG_RED `#921F0B`, CORAL `#F16249`, WINE `#794039`, BLUE `#356CB5` 등 실측 11종.
- **폰트:** 맑은 고딕(latin/ea 동시 지정으로 한글 깨짐 방지).
- **좌표 상수:** 롯데알미늄 `_final` PPTX XML 전수 추출 실측값(`core/designer.py`에 Python 상수로 보존).
- **도식 패턴:** TOC / Pain Point 2컬럼 / 프로세스 / diff_matrix / approach_vs / 컨테이너 등 19종.

---

## 테스트

```bash
cd C:\Users\cgpar\ppt-skill
python -m pytest -v
```

- `test_schema` — 19종 타입 레지스트리, 알 수 없는 타입/누락 필드 검출.
- `test_planner` — config→spec 구조화, 섹션 로마숫자, 입력 불변, pptx 비의존.
- `test_cli` — `--list/--validate/--dry-run` 및 친절한 에러.
- `test_integration` — 샘플 config 엔드투엔드. 전체 렌더는 실제 템플릿 존재 시에만 수행(없으면 skip).
- `test_designer_smoke` — 엔진 API 보존 + 디자인 토큰 소비.

---

## 분석 기반 / 연혁

- **v1.0** — 10개 제안서 530+ 파일 분석으로 통합 스킬(기획+디자인) 등록.
- **v2.0~v4.0** — 디자인/좌표 정밀화(롯데알미늄 `_final` XML 전수 추출), 기아·파라다이스 _final 팔레트 역추출.
- **스킬 분리** — 통합 스킬을 **기획/디자인 2개 스킬 × JSON/MD 2포맷**으로 분리.
- **통합 프레임워크 (2026-06)** — 프로젝트별 명령형 스크립트를 데이터 구동(`config/*.json` + `core/` + `main.py`)으로
  리팩터링. 엔진은 `core/designer.py`로 보존, 레거시는 `archive_history/`로 동결.

> 최우선 규칙: `fb_`(리더 피드백) · `_final`(최종본) 기준이 초안과 충돌 시 항상 `fb_`/`_final` 우선.
