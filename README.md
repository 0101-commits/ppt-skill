# ppt-skill — HCG 제안서 PPT 자동화 스킬

휴먼컨설팅그룹(HCG) 컨설팅 제안서를 **기획(Planning)** 과 **디자인(Design)** 두 스킬로 분리해 관리하고, `auto_ppt.py`로 실제 PPTX를 자동 생성하는 저장소.

10개 제안서 약 530개 파일(피드백 `fb_`, 최종본 `_final`, 내부회의록 포함)을 교차 분석해 도출한 HCG 제안서 표준을 코드화했다.

---

## 핵심 구조: 스킬 = 2개 × 2포맷

스킬은 **기획/디자인 2개**로 나뉘고, 각각 **JSON과 MD 두 버전을 별도 관리**한다.

| 스킬 | JSON (프롬프트/기계 판독) | MD (사람 판독/리뷰) | 담당 영역 |
|------|--------------------------|---------------------|-----------|
| **기획 (Planning)** | `skill_ppt_planning.json` | `skill_ppt_planning.md` | 스토리라인, 목차, 헤드 메시지, 논리 흐름, 맥락 요약 |
| **디자인 (Design)** | `skill_ppt_design.json` | `skill_ppt_design.md` | 톤앤매너, 색상/폰트, 레이아웃, 좌표 상수, 도식화 패턴 |

> **JSON vs MD:** 동일 내용을 두 포맷으로 **병행 유지**한다.
> - **JSON** — LLM 프롬프트 주입·자동화 파이프라인용. 키 기반 구조.
> - **MD** — 사람이 읽고 리뷰·수정하는 정본. 표/코드블록/체크리스트.
>
> 한쪽을 수정하면 다른 쪽도 함께 갱신해 두 버전을 동기 상태로 유지한다.

---

## 파일 맵

```
ppt-skill/
├── README.md                    # (이 파일)
├── skill_ppt_planning.json      # 기획 스킬 — JSON
├── skill_ppt_planning.md        # 기획 스킬 — MD
├── skill_ppt_design.json        # 디자인 스킬 — JSON
├── skill_ppt_design.md          # 디자인 스킬 — MD (구 HCG_PPT_Design_Skill.md, v4.0)
├── auto_ppt.py                  # python-pptx 자동 생성기 (디자인 스킬 좌표 구현)
└── HCG_Automated_Draft.pptx     # auto_ppt.py 출력 샘플
```

---

## 두 스킬 요약

### 1. 기획 (`skill_ppt_planning`)
- **표준 스토리라인:** 표지 → 목차 → `Ⅰ 프로젝트 배경` → `Ⅱ 추진방안(3모듈)` → `Ⅲ 일정·조직` → `Ⅳ HCG 소개` → End
- **핵심 메시지 원칙:** 제도(MBO/OKR/직무급)는 껍질, 본질은 **평가-보상 연계 구조**. 한 문장 주장으로 수렴.
- **헤드 메시지:** 결론 선행 · 중립성 · 전체 통합 · 오해 방지.
- **배경 전개:** `변화 → 현상 → 가설 이슈 → HCG 관점 → 대응` (문제 나열식 금지).
- **표준 분량:** 17~24장(고객 요청 우선, 대형 RFP는 80~100장 허용).
- 기존 design skill **md v1.0의 기획 영역을 계승·분리**.

### 2. 디자인 (`skill_ppt_design`)
- **파일 스펙:** 10.833" × 7.5", 레이아웃 4종(표지/목차/본문/End), 흰 배경.
- **테마 색:** HCG_RED `#921F0B`, CORAL `#F16249`, WINE `#794039`, GRAY 계열, BLUE `#356CB5` 등 실측 11종.
- **폰트:** 맑은 고딕(latin/ea 동시 지정으로 한글 깨짐 방지).
- **좌표 상수:** 롯데알미늄 `_final` PPTX XML 전수 추출 실측값.
- **도식 패턴:** TOC / Pain Point 2컬럼 / 프로세스 / 단일컬럼 / 참고.

---

## 사용법

### 스킬 로드 순서
신규 제안서 착수 시 **기획 → 디자인** 순으로 로드한다.
1. `skill_ppt_planning` 으로 목차·헤드 메시지·슬라이드 타입 태그 확정.
2. 확정 산출물을 `skill_ppt_design` 으로 핸드오프 → 좌표·색상·레이아웃 매핑.

### 자동 생성
```bash
cd C:\Users\cgpar\ppt-skill
python auto_ppt.py        # → HCG_Automated_Draft.pptx
```
`auto_ppt.py`는 실제 `_final` PPTX를 템플릿으로 열어 테마/폰트/레이아웃을 100% 상속한 뒤 슬라이드를 재생성한다.

**의존성:** Python 3.10+, `python-pptx`

```bash
pip install python-pptx
```

---

## 핸드오프 흐름

```
[skill_ppt_planning]  목차 + 헤드 메시지 + 타입 태그
        │  (cover/toc/overview/painpoint2col/process/single/2col/reference/end)
        ▼
[skill_ppt_design]    타입 태그 → 좌표/색상/레이아웃 매핑
        │
        ▼
[auto_ppt.py]         실제 PPTX 생성
```

---

## 분석 기반 / 연혁

- **v1.0** — 10개 제안서 530+ 파일 분석으로 통합 스킬(기획+디자인) 등록.
- **v2.0~v4.0** — 디자인/좌표 정밀화(롯데알미늄 `_final` XML 전수 추출).
- **분리** — 통합 스킬을 **기획/디자인 2개 스킬 × JSON/MD 2포맷**으로 분리, 기획은 v1.0 기획 영역 계승.

> 최우선 규칙: `fb_`(리더 피드백) · `_final`(최종본) 기준이 초안과 충돌 시 항상 `fb_`/`_final` 우선.
