# -*- coding: utf-8 -*-
"""
HCG 롯데케미칼 직무기반 HR 제도 설계 및 도입 제안서 (Draft)
- 기획: skill_ppt_planning (스토리라인 Ⅰ배경→Ⅱ추진방안 3모듈→Ⅲ일정·조직→Ⅳ HCG소개)
- 디자인: skill_ppt_design (롯데알미늄 _final 템플릿 테마 상속, 좌표/색상 실측)
- auto_ppt.py 헬퍼/제너릭 빌더 재사용
실행: python lotte_chemical_ppt.py
"""
import os, sys
sys.stdout.reconfigure(encoding='utf-8')

# auto_ppt.py 의 검증된 헬퍼/제너릭 빌더/상수 재사용
from auto_ppt import (
    Presentation, REAL, Inches, PP_ALIGN, MSO_AUTO_SHAPE_TYPE, qn, etree,
    HCG_RED, DARK_GRAY, MED_GRAY, WHITE, BLACK,
    LABEL_X, LABEL_W, COL_L_X, COL_R_X, ITEM_W, ITEM_Y0, ITEM_DY, HDR_Y,
    delete_all_slides, add_item, add_textbox, set_title,
    build_body_2col, build_body_single, build_body_process,
    build_section_divider, build_end,
)

OUT = "C:\\Users\\cgpar\\ppt-skill\\HCG_롯데케미칼_직무기반HR_제안서_Draft.pptx"


# ── 표지 ───────────────────────────────────────────────────
def build_cover(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:
            ph.text = "롯데케미칼 직무기반 HR제도 설계 및 도입"
        elif idx == 1:
            ph.text = "- 제안서 -"
    add_textbox(slide, "2026.06", 1.137, 5.319, 8.563, 0.30,
                fsize=10, align=PP_ALIGN.CENTER, color=DARK_GRAY)


# ── 목차 (대챕터 4개) ──────────────────────────────────────
def build_toc(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    add_textbox(slide, "목차 (Contents)", 0.5, 0.3, 9.8, 0.5,
                fsize=18, bold=True, color=HCG_RED)
    line = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE,
                                  Inches(0.5), Inches(0.9), Inches(9.8), Inches(0.03))
    line.fill.solid(); line.fill.fore_color.rgb = HCG_RED
    line.line.fill.background()

    toc = [
        ("Ⅰ", "프로젝트 배경", "03-13"),
        ("Ⅱ", "추진방안 (직무체계 / 평가 / 보상)", "14-28"),
        ("Ⅲ", "일정 및 조직", "29-32"),
        ("Ⅳ", "HCG 소개", "33-37"),
    ]
    y = 1.3
    for num, title, pg in toc:
        circle = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL,
                                        Inches(0.6), Inches(y - 0.04), Inches(0.45), Inches(0.45))
        circle.fill.solid(); circle.fill.fore_color.rgb = HCG_RED
        circle.line.fill.background()
        cp = circle.text_frame.paragraphs[0]; cp.alignment = PP_ALIGN.CENTER
        from auto_ppt import _add_run
        _add_run(cp, num, fsize=14, bold=True, color=WHITE)
        add_textbox(slide, title, 1.3, y, 7.3, 0.5, fsize=15, color=DARK_GRAY)
        add_textbox(slide, pg, 8.8, y, 1.2, 0.5, fsize=12, align=PP_ALIGN.RIGHT, color=MED_GRAY)
        y += 1.0


# ── Project Overview ───────────────────────────────────────
def build_overview(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    set_title(slide, "Project Overview",
              "롯데케미칼의 직무체계 표준화 및 직무기반 평가/보상제도 운영방안 도출")

    add_textbox(slide, "추진 내용", LABEL_X, 1.70, LABEL_W, 0.32,
                fsize=12, italic=True, align=PP_ALIGN.RIGHT, color=MED_GRAY)
    add_textbox(slide, "다(多)사업장·M&A 통합 이력으로 분산된 직무 데이터를 글로벌 표준+AI로 신속 표준화하고, 직군 특성을 반영한 평가/보상제도를 설계하여 직무기반 HR제도를 도입",
                1.836, 1.70, 8.759, 0.45, fsize=11, color=DARK_GRAY)

    add_textbox(slide, "추진 프로세스", 0.0, 2.55, 1.62, 0.32,
                fsize=12, italic=True, align=PP_ALIGN.RIGHT, color=MED_GRAY)
    steps = [("현황 진단 및\n방향성 수립", 2.133), ("직무체계\n표준화", 4.161),
             ("평가제도\n개선", 6.168), ("보상제도\n개선", 8.174)]
    for title, x in steps:
        add_item(slide, title, x, 2.55, width=1.976, height=0.666,
                 fsize=11, align=PP_ALIGN.CENTER, fill_rgb=HCG_RED, color=WHITE)
    for x in [4.10, 6.10, 8.00]:
        add_textbox(slide, "→", x, 2.78, 0.15, 0.25, fsize=12, align=PP_ALIGN.CENTER, color=MED_GRAY)
    subs = [(2.133, "현황 Quick Review\n방향성 도출"), (4.161, "글로벌 표준+AI\n직무분류·기술서"),
            (6.168, "직군 차별화\n평가모델 설계"), (8.174, "보상 원칙/전략\n재원 Simulation")]
    for x, txt in subs:
        add_textbox(slide, txt, x, 3.25, 1.951, 0.55, fsize=9, color=DARK_GRAY)

    add_textbox(slide, "추진 방안\n(잠정)", LABEL_X, 4.25, LABEL_W, 0.50,
                fsize=12, italic=True, align=PP_ALIGN.RIGHT, color=MED_GRAY)
    plan = ("추진 기간 : 16주 (직무체계 표준 초안 약 6주 + Fine-tuning·평가/보상 병행)\n"
            "투입 인력 : PM 1명 + 직무/평가·보상 전문 컨설턴트 2~3명\n"
            "추진 금액 : 별도 협의 (범위·투입형태에 따라 변동; 부가세 별도)")
    add_textbox(slide, plan, 1.836, 4.25, 8.759, 1.10, fsize=10, color=DARK_GRAY)


def main():
    print("Loading template:", REAL)
    prs = Presentation(REAL)
    delete_all_slides(prs)
    n = [0]
    def mark(label):
        n[0] += 1; print(f"  S{n[0]:02d} {label}")

    # ── Ⅰ. 프로젝트 배경 ──────────────────────────────────
    build_cover(prs); mark("표지")
    build_toc(prs); mark("목차")
    build_overview(prs); mark("Project Overview")

    build_body_single(prs, "산업에 대한 이해",
        "석유화학 산업의 전환기 — 직무 중심 HR이 인력 운영 효율의 핵심 변수로 부상",
        [("산업 환경", "글로벌 공급과잉·탄소중립 전환·고부가 스페셜티 확대 → 인력 재배치·전문성 요구 증대"),
         ("HR 화두", "롯데그룹 지주 차원의 '직무 중심 HR' 전환 요구 — 계열사 공통 과제"),
         ("생산 현장", "교대·공정 기반 생산직 비중 높아 사무직 중심 제도로는 직무가치 차등화 한계"),
         ("시사점", "직무를 표준 언어로 정의해야 평가·보상·배치의 일관성과 공정성 확보 가능")])
    mark("산업 이해")

    build_body_single(prs, "롯데케미칼에 대한 이해",
        "다(多)사업장·M&A 성장 이력이 만든 직무체계의 분절 — 표준화가 선결 과제",
        [("사업 구조", "기초소재·첨단소재·정밀화학 등 다수 사업부 / 국내외 복수 사업장 운영"),
         ("성장 이력", "인수합병을 통한 확장으로 사업장별 일하는 방식·R&R·직무 정의가 상이"),
         ("핵심 이슈", "직무·조직·사람이 분리되지 않아 '표준 직무'가 부재 → 조직마다 기준 상이"),
         ("지주 가이드", "지주 표준 직무 가이드는 사무직 중심 — 생산직 적용 기준·구체 설명 미흡")])
    mark("롯데케미칼 이해")

    # Pain Point (As-Is) ↔ Needs
    build_body_2col(prs, "Client's Needs & Pain Point",
        "직무체계·평가/보상 미흡에 따른 인사 개편이 필요하나, 전통적 컨설팅은 자원·기간 과다",
        "현재 문제점 (As-Is)", "기대 개선사항 (To-Be)",
        ["표준 직무 부재 — 사업장별 직무 정의·R&R 상이",
         "직무 데이터 노후·분산 — 현행화·활용 곤란",
         "생산직 평가 기준 미흡 — 직군 특성 미반영",
         "보상 차등화 근거 부족 — 직무가치 미반영",
         "전통적 컨설팅 — Interview/Survey로 구성원 부담·장기 소요"],
        ["글로벌 표준+AI로 직무분류체계 신속 표준화",
         "AI 직무기술서로 전(全)직무 현행화",
         "직군별 차별화 평가모델 (성과/투입/역량 지향)",
         "직무가치 기반 Pay 정책선·차등 설계",
         "표준 설계도 기반 — 시간/비용/부담 최소화"])
    mark("Pain Point")

    # HCG Approach (Legacy vs HCG)
    build_body_2col(prs, "HCG's Approach",
        "HR 전문가 지식/경험 + 글로벌 표준을 집대성한 '표준 설계도'와 AI로 Pain Point를 쉽고·정확하고·빠르게 해소",
        "Legacy Consulting", "HCG Approach (표준 설계도 + AI)",
        ["맞춤 구축 — 시간/비용/복잡성 ↑",
         "대량 리소스·구성원 업무조사 부담",
         "산출물 간 低연계·사후관리 부족",
         "글로벌 표준 vs 국내 현장 간 Gap"],
        ["국제표준(APQC·ESCO·ISCO·O*NET·NCS) 기반 빠른 초안",
         "AI 직무기술서 — 조사 부담 최소화",
         "직무→평가→보상 일관 연계 설계",
         "지주 가이드·현장 맥락을 Custom으로 정합"])
    mark("HCG Approach")

    build_body_single(prs, "[참고] 高맥락·AI 기반 컨설팅 접근",
        "표면 문맥을 넘어 고객 고유의 비언어적 맥락(HR Gene)을 이해하고 HR 모듈에 즉시 반영",
        [("맥락 이해", "사업장별 공정·교대·합의 배경 등 비정형 맥락을 직무 표준에 반영"),
         ("AI Agent", "HR 전문가 + AI Agent 협업으로 맥락 중심 직무·역량 도출"),
         ("자산 활용", "JobSkillEngine(자체 플랫폼)으로 산업 표준 직무·스킬 자동 생성"),
         ("효과", "공정성/편향 제거 측면에서 구성원 수용성 제고 — AI 활용 목표·평가 보조")])
    mark("[참고] 高맥락·AI 접근")

    build_body_single(prs, "프로젝트 목적 및 범위",
        "직무체계 표준화를 기반으로 평가·보상까지 직무 중심으로 연계하는 HR제도 도입",
        [("목적", "롯데케미칼 직무 중심 HR 전환 — 표준 직무체계 확립 및 직무기반 평가/보상 운영방안 도출"),
         ("범위 (직무)", "직군→직렬→직무 3단계 표준 분류체계 + 전직무 직무기술서(AI 초안 기반)"),
         ("범위 (평가)", "직군 특성 반영 평가모델 설계 + 평가 운영체계 / 수립 Process Map"),
         ("범위 (보상)", "Pay Policy·Pay Mix·Pay Level·평가-보상 연계 + 재원 Simulation"),
         ("대상", "사무직·생산직 전 직군 (임원 제외) — 사업장 공통 표준 + 사업장 Custom")])
    mark("목적 및 범위")

    # Why HCG
    build_body_2col(prs, "Why HCG",
        "직무기반 HR제도 설계 전문 역량 + AI 자산 + A-to-Z 실행 경험을 모두 보유",
        "차별화 축", "핵심 내용",
        ["1. 유사 프로젝트 경험", "2. 직무·스킬 AI 자산", "3. 제도수립~정착 A-to-Z"],
        ["대기업 제조·화학 직무체계/평가/보상 다수 구축",
         "JobSkillEngine — 산업별 표준 직무·스킬 자동 생성",
         "설계뿐 아니라 도입·코칭·정착까지 End-to-End 지원"])
    mark("Why HCG")

    build_body_single(prs, "Why HCG  1. 유사 프로젝트 수행 경험",
        "대기업 제조·화학 산업의 직무기반 HR / 평가·보상 프로젝트 다수 수행",
        [("직무체계", "롯데알미늄 직무기반 HR, 코오롱인더스트리 직무분석, 포스코인터내셔널 HR 플랫폼 등"),
         ("평가·보상", "기아·코오롱·GS에너지·녹십자 등 평가·보상체계 설계 및 운영 지원"),
         ("화학·제조", "석유화학·소재·제조업 동종/유사 산업 직무 및 생산직 보상 수행 사례"),
         ("그룹사", "롯데 계열 직무 중심 HR 전환 맥락에 대한 높은 이해도")])
    mark("Why HCG 1")

    build_body_single(prs, "Why HCG  2. 직무·스킬 AI 자산 (JobSkillEngine)",
        "산업별 표준 직무분류·직무기술서·스킬 디렉토리를 AI로 자동 생성하는 자체 플랫폼 보유",
        [("Module 1", "직무분류체계 — 가치사슬 분석 기반 직군→직렬→직무 3단계 자동 생성"),
         ("Module 2", "직무기술서 — 미션·과업·스킬(KSAO)·자격요건 자동 생성"),
         ("Module 4", "스킬 디렉토리 — 스킬 정규화·관계 분석(AI/PMI/벡터/온톨로지 4지표)"),
         ("참조 표준", "APQC·ESCO·ISCO·O*NET·NCS 등 국제/국내 표준 직무 데이터 내장")])
    mark("Why HCG 2")

    build_body_single(prs, "Why HCG  3. 제도수립~정착 'A to Z' 경험",
        "제도 설계에 그치지 않고 도입·교육·코칭·시스템 연계까지 정착을 책임",
        [("설계", "직무·평가·보상 제도 설계 및 사업장별 Custom"),
         ("도입/코칭", "리더·구성원 대상 변화관리, 평가자 교육, 운영 코칭"),
         ("시스템", "talenX 등 목표수립/리뷰 지원 도구 연계로 운영 효율화"),
         ("정착", "운영 가이드라인 + 주기적 업데이트로 HR Data 축적·고도화")])
    mark("Why HCG 3")

    # ── Ⅱ. 추진방안 ───────────────────────────────────────
    build_section_divider(prs, "Ⅱ", "추진방안",
                          "직무체계 표준화 → 직군 차별화 평가 → 직무기반 보상의 3개 모듈 추진")
    mark("Section Ⅱ")

    build_body_process(prs, "Overall Approach",
        "현황진단을 토대로 3개 모듈을 단계적·병행 추진하며 변화관리를 동반",
        [("현황 진단", "Quick Review\n방향성 수립"),
         ("M1 직무체계", "표준화\nAI 직무기술서"),
         ("M2 평가제도", "직군 차별화\n평가모델"),
         ("M3 보상제도", "Pay 설계\nSimulation")],
        [("변화관리", "전 과정에 리더·구성원 커뮤니케이션 및 코칭 동반 진행"),
         ("산출물", "표준 직무체계(안), 직무기술서, 평가/보상제도(안), 운영 가이드라인")])
    mark("Overall Approach")

    build_body_process(prs, "M1. 직무체계 표준화  Overview",
        "지주 가이드와 현장 맥락을 글로벌 표준+AI로 통합해 신속하게 표준 직무를 확립",
        [("지주 가이드\n해석", "표준 직무\n원칙 decipher"),
         ("글로벌 BM\n+AI 수집", "국제표준 직무\n체계 수집"),
         ("표준 직무\n도출", "사업장 공통\n표준 확립"),
         ("사업장\nCustom", "공정·교대 맥락\n반영 조정")])
    mark("M1 Overview")

    build_body_process(prs, "직무분류체계 Framework",
        "석유화학 직무 특성을 반영한 직군→직렬→직무 3단계 표준 분류체계",
        [("직군\n(Job Family)", "생산, 기술/연구,\n영업, 경영지원 등"),
         ("직렬\n(Job Series)", "직군 내\n유사 직무 묶음"),
         ("직무\n(Job)", "표준 직무 단위\n정의"),
         ("직무기술서\n(JD)", "미션·과업·스킬\n자격요건")],
        [("분류 기준", "업무 유사성·역량 공통성·가치사슬·조직구조 반영"),
         ("참조 표준", "APQC·ESCO·ISCO·O*NET·NCS 국제/국내 표준 매핑")])
    mark("직무분류 Framework")

    build_body_process(prs, "직무체계 표준화 Process",
        "Standard를 빠르게 제시하고 사업장 특수성에 맞추는 4단계 — 1개월 내 표준 초안 도출 목표",
        [("Step1\n표준 수집", "글로벌/국내\n표준 직무 AI 수집"),
         ("Step2\n표준 도출", "지주 가이드 정합\n공통 표준 확립"),
         ("Step3\nCustom", "사업장별 맥락\n자료 수집·조정"),
         ("Step4\n검증", "TFT 협의\n직무 확정")])
    mark("직무체계 표준화 Process")

    build_body_single(prs, "AI 활용 직무기술서 작성",
        "JobSkillEngine으로 직무기술서 초안을 자동 생성하고 현업 검증으로 현행화",
        [("작성 방법", "AI 초안 생성 → 현업 인터뷰/검증 → 확정 (조사 부담 최소화)"),
         ("작성 항목", "직무 미션, 핵심 과업/활동, 스킬(KSAO), 자격요건, 평가지표 연계"),
         ("활용 도구", "JobSkillEngine(Claude API 기반) + HCG 표준 JD 템플릿"),
         ("품질 관리", "직무 담당자·팀장 검토 → HR 최종 확인의 3단계 검증")])
    mark("AI 직무기술서")

    build_body_single(prs, "[참고] 직무기술서 구성 체계 (KSAO)",
        "표준 JD 양식 — 직무 수행에 필요한 지식·스킬·능력·기타특성 정의",
        [("Knowledge", "직무 수행에 필요한 전문 지식 영역 (공정·소재·안전·규제 등)"),
         ("Skill", "실무 수행 스킬 — 정규화된 스킬 디렉토리와 연계"),
         ("Ability", "직무 요구 능력/역량 (문제해결·협업·전문성 등)"),
         ("Other", "자격증·경력·교대 가능 여부 등 직무 특이 요건")])
    mark("[참고] KSAO")

    build_body_2col(prs, "M2. 평가제도 개선  현황 및 방향",
        "사무직 중심 평가를 직군 특성에 맞게 차별화 — 생산직 적용 가능한 모델 필요",
        "현행 문제점", "개선 방향",
        ["단일 평가틀 — 직군 특성 미반영",
         "생산직 평가 기준 부재",
         "목표·KPI 연계 형식화",
         "평가 결과 수용성 저하"],
        ["직군별 차별화 평가모델 설계",
         "투입·역량 지향 등 생산직 적합 모델",
         "직무 KPI 기반 목표설정 체계",
         "AI 보조로 공정성·편향 완화"])
    mark("M2 평가 현황")

    build_body_process(prs, "직군별 차별화 평가모델",
        "직무 특성을 도출해 일의 성격에 따라 평가 방식을 차등 적용",
        [("성과 지향", "정량 KPI 중심\n(영업·사업)"),
         ("투입 지향", "공정·시간 기반\n(생산직)"),
         ("역량/전문", "전문성·Potential\n(연구·기술)"),
         ("직군 적용", "직군/직렬별\n평가 틀 매핑")],
        [("도출 방식", "직무 확정 후 일의 특성을 분석해 평가 유형 분류"),
         ("협의", "TFT 협의로 직군별 적합 평가제도 수립")])
    mark("직군 차별화 평가")

    build_body_process(prs, "평가제도 수립 Process Map",
        "가설적 결론이 아닌 '도출 과정'을 보여주는 4단계 수립 프로세스",
        [("1. 직무 특성\n도출", "직무별 일의\n성격 분석"),
         ("2. 평가 유형\n분류", "성과/투입/역량\n지향 구분"),
         ("3. 평가 틀\n설계", "직군별 KPI·\n역량 지표"),
         ("4. 운영체계\n수립", "주기·등급·\n캘리브레이션")])
    mark("평가 Process Map")

    build_body_single(prs, "[참고] AI 기반 목표추천 · 업무 Check-in",
        "talenX 등 시스템으로 목표 수립·리뷰를 보조해 운영 효율화",
        [("목표 추천", "직무·과거 데이터 기반 AI 목표 추천으로 목표설정 부담 완화"),
         ("업무 Check-in", "상시 업무 기록·리뷰로 평가 근거 축적 (포스코·두산式 실패 방지)"),
         ("효율화", "관리자 부담 최소화 — 시스템 자동 트래킹으로 passive 운영"),
         ("연계", "수립된 직무·평가 데이터와 연동")])
    mark("[참고] AI 목표·Check-in")

    build_body_process(prs, "평가 운영체계",
        "연간 성과관리 Cycle — 목표설정부터 보상 연계까지 일원화",
        [("연초\n목표설정", "직무 KPI\n목표 합의"),
         ("반기\n중간점검", "실적 Review\n코칭·피드백"),
         ("연말\n종합평가", "성과+역량\n종합 평가"),
         ("차년\n결과활용", "보상 연계\n육성 계획")])
    mark("평가 운영체계")

    build_body_2col(prs, "M3. 보상제도 개선  원칙 및 방향",
        "직무가치와 직군 특성을 반영한 보상 — '어떤 순서로 개선할지' 프로세스 중심",
        "현행 보상 구조", "개선 방향",
        ["연공·획일 — 직무가치 미반영",
         "직군 간 보상 형평성 이슈",
         "성과 연계 약함",
         "외부 경쟁력 검증 미흡"],
        ["Pay Policy(보상 원칙/정책선) 정립",
         "사무직·생산직 Pay Mix 차별화",
         "평가-보상 Contributor Connection",
         "시장 BM 기반 Pay Level 결정"])
    mark("M3 보상 원칙")

    build_body_single(prs, "직무기반 보상제도 설계",
        "Pay Policy → Pay Mix → Pay Level → 평가 연계의 순차적 보상 개선 프로세스",
        [("Pay Policy", "보상 원칙·정책선(시장 대비 지향 수준) 가이드 수립"),
         ("Pay Mix", "사무직·생산직 등 직군별 기본급/변동급 구성 설계"),
         ("Pay Level", "충원 필요 직군 중심 시장 BM으로 외부 경쟁력 확보"),
         ("평가 연계", "평가 결과를 기본급/변동급에 연동 (Contributor Connection)")])
    mark("보상제도 설계")

    build_body_single(prs, "보상 Simulation",
        "변동급 상세 설계보다 재원 영향 Simulation으로 의사결정 지원",
        [("목표", "현행 인건비 영향 범위 내에서 직무가치 반영 구조 전환 검토"),
         ("시나리오", "현행 유지형 / 점진 전환형 / 즉시 전환형 비교"),
         ("분석", "직무등급 × Pay Band 대비 현행 급여 분포 분석"),
         ("협의", "노조·구성원 커뮤니케이션 및 이해관계자 합의 계획 포함")])
    mark("보상 Simulation")

    # ── Ⅲ. 일정 및 조직 ──────────────────────────────────
    build_section_divider(prs, "Ⅲ", "일정 및 조직",
                          "16주 추진 일정, 수행조직, 투입인력 및 제안 비용")
    mark("Section Ⅲ")

    build_body_process(prs, "추진 일정 (16주)",
        "직무체계 표준 초안을 빠르게 확보하고 Fine-tuning·평가/보상을 병행",
        [("W1-2\n킥오프", "현황자료 수집\n방향성 수립"),
         ("W3-6\n직무표준", "표준 직무체계\n초안 도출"),
         ("W7-10\n직무확정", "사업장 Custom\nAI 직무기술서"),
         ("W11-14\n평가·보상", "평가모델·보상\n설계·Simulation"),
         ("W15-16\n완료", "운영 가이드\n최종 보고")])
    mark("추진 일정")

    build_body_2col(prs, "수행 조직 및 역할",
        "롯데케미칼 TFT와 HCG 전문 컨설턴트의 긴밀한 협업 체계",
        "롯데케미칼", "HCG",
        ["프로젝트 스폰서 / HR 주관",
         "직무 TFT (사업장·직군 대표)",
         "현업 인터뷰/검증 참여",
         "의사결정 및 내부 합의"],
        ["PM — 전체 총괄·품질 관리",
         "직무체계 전문 컨설턴트",
         "평가·보상 전문 컨설턴트",
         "JobSkillEngine AI 운영 지원"])
    mark("수행 조직")

    build_body_single(prs, "투입 인력 및 제안 비용",
        "롯데케미칼 직무기반 HR제도 설계 및 도입 컨설팅 (잠정 — 협의 변동)",
        [("인력 구성", "PM 1명 + 직무체계 전문 1명 + 평가·보상 전문 1~2명"),
         ("투입 기간", "16주 (약 4개월) — 킥오프~최종 보고"),
         ("제안 금액", "별도 협의 (범위·투입형태·사업장 수에 따라 변동; 부가세 별도)"),
         ("산출물", "표준 직무체계(안), 직무기술서(전직무), 평가/보상제도(안), 운영 가이드라인")])
    mark("투입 인력·비용")

    # ── Ⅳ. HCG 소개 ──────────────────────────────────────
    build_section_divider(prs, "Ⅳ", "HCG 소개",
                          "국내 최대·유일의 HR End-to-End 전문가 그룹")
    mark("Section Ⅳ")

    build_body_single(prs, "HCG 소개  Overview",
        "휴먼컨설팅그룹(HCG)은 HR 컨설팅부터 시스템·AI까지 End-to-End 서비스를 제공",
        [("정체성", "국내 최대·유일의 HR 전문가 그룹 — 컨설팅 + Implementation + 시스템"),
         ("전문성", "직무·평가·보상·조직문화·HR DT 전 영역 컨설팅 역량"),
         ("AI 역량", "JobSkillEngine·talenX 등 자체 HR AI/솔루션 보유"),
         ("실적", "870여 고객사 / 다수 대기업 HR 제도 구축 경험")])
    mark("HCG Overview")

    build_body_2col(prs, "HCG 사업 영역 / 컨설팅 서비스 영역",
        "HR 전 생애주기를 아우르는 컨설팅·솔루션 포트폴리오",
        "사업 영역", "컨설팅 서비스 영역",
        ["HR 컨설팅", "HR 솔루션/시스템", "HR AI (직무·스킬)", "교육·코칭"],
        ["직무체계·직무분석", "평가·보상제도 설계", "조직문화·직원경험", "성과관리(OKR)·HR DT"])
    mark("사업/컨설팅 영역")

    build_body_single(prs, "주요 고객사 및 Long-term Relationship",
        "대기업·제조·화학·금융 등 전 산업 870여 고객사와의 장기 파트너십",
        [("제조/화학", "롯데알미늄, 코오롱인더스트리, 포스코인터내셔널, 코스맥스 등"),
         ("대기업", "기아, 현대자동차, GS에너지, 매일그룹 등"),
         ("바이오/제약", "녹십자, 한림제약, 네오이뮨텍 등"),
         ("파트너십", "단발 프로젝트가 아닌 제도 정착까지의 Long-term Relationship 지향")])
    mark("주요 고객사")

    build_end(prs); mark("End")

    print(f"\nSaving → {OUT}")
    prs.save(OUT)
    print("Done! 총", n[0], "장")
    return OUT


if __name__ == "__main__":
    main()
