# -*- coding: utf-8 -*-
"""
HCG 기아 중장기 보상체계 개선 제안서 (AI Draft v2 — 인간 _final 밀도 정합)
- 스킬 v1.2 / 디자인 v4.2 적용: 40~60장 스케일, 미니리포트 밀도(350~500자),
  도식≥1/장(표·로드맵·연결선·아이콘), Why HCG 다장 전개, [참고] 적극분리,
  고유 프레임워크(Benchmarking Connect / Global Big Data Analytics / HCG HR Maturity / Process Overview)
- RFP + 회의록 6건(09.05 fb 1순위) 반영
실행: python auto_ppt_kia.py
"""
import os, sys
sys.stdout.reconfigure(encoding='utf-8')

from auto_ppt import (
    Presentation, REAL, Inches, PP_ALIGN, MSO_AUTO_SHAPE_TYPE,
    HCG_RED, DARK_GRAY, MED_GRAY, WHITE, BLACK, GRAY_LT, BLUE,
    LABEL_X, LABEL_W,
    delete_all_slides, add_item, add_textbox, set_title, _add_run,
    build_body_2col, build_body_single, build_body_process,
    build_section_divider, build_end,
    build_overview_3col, build_diff_matrix, build_pain_point_categorized,
    build_approach_vs, add_insight_quote, add_example_badge,
    # v2 (v4.2)
    add_connector_line, add_arrow_flow, add_icon_placeholder, add_table,
    build_process_roadmap, build_compare_table,
)

OUT = ("C:\\Users\\cgpar\\OneDrive - 휴먼컨설팅그룹\\09 Admin\\09 etc\\other\\Claude\\"
       "기아 제안서\\HCG_기아_보상체계_개선_제안서_AI_Draft.pptx")


def build_cover(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0: ph.text = "기아 중장기 보상체계 개선 추진"
        elif idx == 1: ph.text = "- 제안서 -"
    add_textbox(slide, "2025.09", 1.137, 5.319, 8.563, 0.30, fsize=10, align=PP_ALIGN.CENTER, color=DARK_GRAY)


def build_toc(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    add_textbox(slide, "목차 (Contents)", 0.5, 0.3, 9.8, 0.5, fsize=18, bold=True, color=HCG_RED)
    ln = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0.5), Inches(0.9), Inches(9.8), Inches(0.03))
    ln.fill.solid(); ln.fill.fore_color.rgb = HCG_RED; ln.line.fill.background()
    toc = [("Ⅰ", "제안 배경 및 Why HCG", "03-17"),
           ("Ⅱ", "수행 방안 (벤치마킹 / 글로벌 리서치 / 분석·Report)", "18-38"),
           ("Ⅲ", "추진 일정 · 조직 · 비용", "39-43"),
           ("Ⅳ", "HCG 소개", "44-46")]
    y = 1.3
    for num, title, pg in toc:
        c = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(0.6), Inches(y - 0.04), Inches(0.45), Inches(0.45))
        c.fill.solid(); c.fill.fore_color.rgb = HCG_RED; c.line.fill.background()
        cp = c.text_frame.paragraphs[0]; cp.alignment = PP_ALIGN.CENTER
        _add_run(cp, num, fsize=14, bold=True, color=WHITE)
        add_textbox(slide, title, 1.3, y, 7.3, 0.5, fsize=14, color=DARK_GRAY)
        add_textbox(slide, pg, 8.8, y, 1.2, 0.5, fsize=12, align=PP_ALIGN.RIGHT, color=MED_GRAY)
        y += 1.0


def build_overview(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    set_title(slide, "Project Overview",
              "객관적 벤치마킹·문헌조사·구성원 의견조사로 기아 보상제도의 現 주소를 진단하고, 근본적 개선의 필요성·당위성을 제시")
    add_textbox(slide, "추진 배경", LABEL_X, 1.66, LABEL_W, 0.32, fsize=12, italic=True, align=PP_ALIGN.RIGHT, color=MED_GRAY)
    add_textbox(slide, "연공급(호봉) 중심·획일적 보상으로 인건비 지속 상승 / 시장 임금 수준·Trend 정보 부족 / 통상임금·평균임금 등 法·정책 Risk로 추가 인건비 상승 예상",
                1.836, 1.66, 8.759, 0.55, fsize=11, color=DARK_GRAY)
    add_textbox(slide, "추진 범위", 0.0, 2.58, 1.62, 0.32, fsize=12, italic=True, align=PP_ALIGN.RIGHT, color=MED_GRAY)
    steps = [("타사 벤치마킹", 2.133), ("글로벌 리서치", 4.161), ("구성원 의견", 6.168), ("진단·Report", 8.174)]
    for t, x in steps:
        add_item(slide, t, x, 2.58, width=1.976, height=0.62, fsize=11, align=PP_ALIGN.CENTER, fill_rgb=HCG_RED, color=WHITE, no_border=True)
    add_arrow_flow(slide, [(2.133 + 0.99, 2.89), (4.161 + 0.99, 2.89), (6.168 + 0.99, 2.89), (8.174 + 0.99, 2.89)])
    subs = [(2.133, "글로벌 완성차·\n국내 기업 BM"), (4.161, "Region 환경·\n임금체계 맥락"),
            (6.168, "심층 설문·\n인터뷰"), (8.174, "현황진단·\n시사점 Report")]
    for x, txt in subs:
        add_textbox(slide, txt, x, 3.28, 1.951, 0.55, fsize=9, color=DARK_GRAY)
    add_textbox(slide, "추진 방안\n(잠정)", LABEL_X, 4.30, LABEL_W, 0.50, fsize=12, italic=True, align=PP_ALIGN.RIGHT, color=MED_GRAY)
    plan = ("• 추진 기간 : 12주 (9월말 Kick-off ~ 연말 경영층 보고; 중간보고·결과보고 운영)\n"
            "• 투입 인력 : PM 1명 + 보상/BM 전문 컨설턴트 2명 (가용 6명 중 자동차·대기업 C&B 전문가)\n"
            "• 추진 금액 : 옵션형 — 옵션1 진단 12주 2억 초반대 / 옵션2 후속 설계 포함 가견적 (경비 포함·부가세 별도)\n"
            "• 중점 목표 : 보상제도 '진단' — 개선 방향성 도출·설계는 후속 프로젝트로 추진")
    add_textbox(slide, plan, 1.836, 4.30, 8.759, 1.35, fsize=10, color=DARK_GRAY)
    add_insight_quote(slide, "진단의 목적은 '개선의 당위성'을 객관적 근거로 입증하는 것", x=0.691, y=6.55, w=9.451)


def main():
    print("Loading template:", REAL)
    prs = Presentation(REAL)
    delete_all_slides(prs)
    n = [0]
    def mark(label):
        n[0] += 1; print(f"  S{n[0]:02d} {label}")

    # ════ Ⅰ. 제안 배경 및 Why HCG ════
    build_cover(prs); mark("표지")
    build_toc(prs); mark("목차")
    build_overview(prs); mark("Project Overview")

    s = build_body_single(prs, "제안 배경  보상 환경 변화 (PEST)",
        "정책·경제·사회·기술 환경의 구조적 변화가 기아 보상체계의 전면적 재편 필요성을 동시다발적으로 제기",
        [("Politics  정책/법", "통상임금·평균임금 판결 및 정책 변화 → 수당·상여의 통상임금 산입 확대로 추가 인건비 상승 압박이 구조화"),
         ("Economy  경제", "연공급 기반 고인건비 구조가 지속되는 가운데, 글로벌 경쟁사는 스마트팩토리·생산성 제고로 단위 인건비를 관리"),
         ("Society  사회", "직무 간 보상 차이 관행과 노동가치 재편 흐름 속에서 사무직(매니저급)·생산직 보상구조의 형평성 이슈가 부각"),
         ("Technology  기술/인력", "SW·신직군 확대와 핵심인재 영입·리텐션 경쟁 심화로 인력 구성·보상 비중의 재설계 필요성 증대")])
    add_icon_placeholder(s, 0.30, 1.70, 0.22, HCG_RED); add_icon_placeholder(s, 0.30, 2.40, 0.22, HCG_RED)
    add_icon_placeholder(s, 0.30, 3.10, 0.22, HCG_RED); add_icon_placeholder(s, 0.30, 3.80, 0.22, HCG_RED)
    add_insight_quote(s, "연공·호봉 중심 보상구조의 전면적 재편 필요성이 대두되는 시점 — 객관적 진단이 그 출발점")
    mark("배경 PEST")

    s = build_body_single(prs, "[참고] 통상임금 / 평균임금 Risk",
        "최근 법·정책 변화로 수당·상여의 통상임금 산입 범위가 확대되며 인건비 구조에 직접적 리스크로 작용",
        [("통상임금", "정기·일률·고정 상여 및 각종 수당의 통상임금 산입 확대 → 연장·야간·휴일수당 기준액 동반 상승"),
         ("평균임금", "퇴직급여 산정 기준 상승으로 충당부채·총인건비 증가 압력"),
         ("수당 더블업", "수당 비중이 높은 현 구조에서 통상임금 확대는 '수당이 수당을 키우는' 더블업 리스크로 연결"),
         ("시사점", "수당 중심 임금구조의 구조적 재편 없이는 법·정책 변화에 따른 인건비 상승을 통제하기 어려움")])
    add_example_badge(s)
    mark("[참고] 통상임금 Risk")

    s = build_body_single(prs, "[참고] 글로벌 경쟁사의 생산성 제고 노력",
        "글로벌 완성차·제조 기업은 보상구조 전환과 자동화로 단위 인건비를 관리하며 경쟁력을 확보",
        [("미국", "테슬라·BYD 등 신흥 플레이어의 직고용·성과 연동 보상, 기가팩토리式 생산성 극대화"),
         ("일본", "직능급에서 성과제로의 단계적 전환, 춘투(春鬪) 베이스업 협상 구조"),
         ("독일", "산별노조(IG Metall) 단체협약 기반 임금, 직무·등급 중심 체계"),
         ("시사점", "스마트팩토리로 블루칼라가 대체되는 구조 변화 하에서 고인건비 구조의 지속 가능성에 대한 question")])
    add_example_badge(s)
    mark("[참고] 글로벌 생산성")

    s = build_body_single(prs, "제안 배경  보상체계에 대한 검토 필요성",
        "내부 우려(호봉·정보부족·法 Risk)와 외부 환경 변화가 맞물려 '지금' 객관적 진단이 필요한 시점",
        [("현 구조", "연공급(호봉)·획일적 보상 → 성과·직무가치와의 연계 약화, 인건비 지속 상승"),
         ("정보 격차", "시장 임금 수준·보상 Trend에 대한 객관적 정보 부족으로 의사결정 근거 미흡"),
         ("리스크", "통상임금·평균임금 등 法 Risk가 인건비 상승을 가속"),
         ("결론", "개선 '방향성 설계'에 앞서, 現 주소를 객관적으로 진단하고 개선의 당위성을 입증하는 것이 본 프로젝트의 과제")])
    add_insight_quote(s, "본 프로젝트는 '설계'가 아니라 '진단' — 개선의 방향을 정하기 위한 객관적 근거와 화두를 제시한다")
    mark("검토 필요성")

    build_pain_point_categorized(prs, "Client's Needs & Pain Point",
        "보상제도 진단의 필요성은 크나, 데이터를 '받아 전달'하는 전통적 벤치마킹으로는 실행 가능한 시사점 도출이 어려움",
        "기아 보상제도 現 이슈", "전통적 벤치마킹의 한계",
        [("연공 중심", "연공급(호봉) 중심·획일적 보상으로 인건비 지속 상승", "직접 BM", "Broad 정책·평균 %tile만으로 세부 파악 한계"),
         ("정보 부족", "시장 임금 수준·보상 Trend 정보 부족", "맥락 해석", "데이터만 받아선 제도·변화·이슈 파악 불가"),
         ("법적 Risk", "통상임금·평균임금 등 法 Risk로 인건비 상승", "구조 분석", "수치 비교만으론 구조적 임금 동인 설명 불가"),
         ("직군 혼재", "사무·생산·정비·영업 직군별 보상구조 상이", "직군 차별화", "일반적 4P로는 직군별 보상 포인트 누락"),
         ("핵심인재", "핵심인재 영입·리텐션 경쟁력 불확실", "패키지 해석", "단순 수준 비교로 핵심인재 처우 설계 불가"),
         ("현장직 글로벌", "현장직 보상 글로벌 비교 기준 부재", "Region 맥락", "노동법·노동환경 차이 미반영 비교 위험")],
        summary_left="객관적 진단 부재 → 개선 당위성 입증 곤란",
        summary_right="데이터 나열형 BM → 실행 가능한 시사점 부족")
    mark("Pain Point")

    build_body_2col(prs, "Why HCG",
        "자동차 산업 이해 + 대기업 보상 경험 + 직군별 차별화 노하우 + 직접 BM 역량 + 데이터 분석력을 모두 보유",
        "차별화 축", "핵심 내용",
        ["1. 자동차 산업에 대한 이해", "2. 대기업 보상제도 경험", "3. 직군별 차별화 노하우",
         "4. 직접·심층 BM 역량", "5. 데이터 확보·분석"],
        ["기아·현대차·현대오토에버 등 완성차 생태계 HR 수행",
         "주요 대기업 C&B(보상) 프로젝트 다수",
         "영업·생산·R&D 직군별 보상 포인트 보유",
         "BM 미팅·밋업·네트워크 (파라다이스·SK하이닉스 등)",
         "크롤링·빅데이터로 직군/직무 상대수준 분석"])
    mark("Why HCG")

    s = build_body_single(prs, "Why HCG  1. 자동차 산업에 대한 이해  기아",
        "기아의 직군 구조·노조 환경·보상 이슈에 대한 직접적 이해를 보유",
        [("직군 구조", "일반/연구(사무)·엔지니어(생산)·기술(정비)·오토컨설턴트(영업)의 상이한 보상 특성 이해"),
         ("노조 환경", "책임급(노조 자동탈퇴)·매니저급(노조원) 임금체계 차이와 형평성 이슈 이해"),
         ("성과관리", "기아 성과관리·MS 벤치마킹 수행 경험으로 평가-보상 연계 맥락 이해"),
         ("시사점", "기아 고유 맥락을 아는 컨설턴트가 진단해야 실행 가능한 시사점 도출 가능")])
    add_icon_placeholder(s, 0.30, 1.95, 0.24, HCG_RED)
    mark("Why HCG 1-기아")

    build_body_single(prs, "Why HCG  1. 자동차 산업에 대한 이해  현대차·현대오토에버",
        "완성차 생태계 전반의 보상·직군 프로젝트 경험으로 산업 구조적 이해를 확보",
        [("현대차", "영업 직군 성과보상 체계 등 완성차 보상 프로젝트 수행 경험"),
         ("현대오토에버", "SW·IT 신직군 보상·HR 경험으로 신직군 이해도 보유"),
         ("남양연구소", "R&D·SW 개발직군 경험으로 신직군 보상 특성 이해"),
         ("의의", "완성차 본사~관계사~연구소를 아우르는 산업 생태계 관점의 보상 이해")])
    mark("Why HCG 1-현대차/오토에버")

    build_body_single(prs, "Why HCG  2. 대기업 보상제도 프로젝트 경험",
        "최근 5개년 주요 대기업 C&B(보상) 프로젝트 수행 — 자동차·제조·대기업 보상 전문성",
        [("보상 설계", "기아·GS에너지·녹십자 등 평가·보상체계 설계 및 운영 지원 다수"),
         ("BM 경험", "글로벌·국내 임금구조/수준 벤치마킹 수행 경험"),
         ("전문 인력", "자동차 산업·대기업 C&B 경험 보유 컨설턴트 중심 투입"),
         ("네트워크", "보상 데이터·BM 네트워크 지속 확보·관리")])
    mark("Why HCG 2-대기업 경험")

    build_diff_matrix(prs, "Why HCG  3. 직군별 차별화 방법론 노하우",
        "일반적 4P가 아닌, 직군별 보상 포인트를 알고 접근하는 차별화 노하우 (수행 사례 기반)",
        [{"group": "영업직군", "trait": "성과 변동·개인 기여 중심", "insight": "성과 연동 보상 설계 노하우", "apply": "현대차 등 영업 성과보상 수행"},
         {"group": "생산직군", "trait": "호봉·수당 중심, 노조", "insight": "호봉 전환 설계 노하우", "apply": "D社(두산인프라코어) 호봉 전환 수행"},
         {"group": "R&D직군", "trait": "장기 성과·전문성", "insight": "직군 차별화 보상 노하우", "apply": "K社(코오롱) R&D 보상 수행"},
         {"group": "신직군(SW)", "trait": "시장 직무가치", "insight": "신직군 보상 이해", "apply": "현대오토에버·남양연구소 경험"}],
        bottom_quote="직군별로 보상 Key Question이 다르다 — HCG는 직군별 보상 포인트를 안다",
        example=True)
    mark("Why HCG 3-직군 노하우")

    s = build_body_single(prs, "[참고] 최근 5개년 유사 프로젝트 수행 경험",
        "자동차·제조·대기업 보상(C&B) 및 직군별 차별화 프로젝트 레퍼런스",
        [("자동차/완성차", "기아 성과관리·MS BM, 현대차 영업 보상, 현대오토에버 HR"),
         ("제조/생산직", "두산인프라코어 호봉 전환, 제조 대기업 보상 BM"),
         ("R&D/신직군", "코오롱 R&D 직군 보상, 남양연구소 SW 직군"),
         ("대기업 C&B", "GS에너지·녹십자 등 평가·보상체계 설계")])
    add_example_badge(s)
    mark("[참고] 5개년 경험")

    build_body_single(prs, "Why HCG  4. 직접·심층 벤치마킹 역량 및 네트워크",
        "데이터 서베이 전담팀을 활용한 간접 BM이 아닌, 직접 확인으로 데이터 현재성과 맥락을 확보",
        [("BM 방식", "직접 미팅·밋업·화상으로 제도·이슈·변화 양상까지 심층 확인 (현장방문은 협의 옵션)"),
         ("네트워크", "파라다이스 컨택포인트, SK하이닉스 실리콘밸리 BM 경험 등 직접 채널 보유"),
         ("현재성", "직접 확인으로 보상데이터의 현재성·신뢰성 제고 — Broad 데이터의 한계 보완"),
         ("프로세스", "질문지 설계 → 현업 인터뷰의 차별화된 BM 프로세스")])
    mark("Why HCG 4-BM 역량")

    build_overview_3col(prs, "Why HCG  5. 데이터 확보·관리·활용 전문성",
        "국내/글로벌 보상 플랫폼 데이터와 자체 크롤링으로 직군/직무별 상대 수준을 분석하는 데이터 역량",
        [{"header": "확보 (Sourcing)", "desc": "국내 대형 보상 플랫폼·글로벌 보상 플랫폼 + 자체 크롤링",
          "ai": "데이터 크롤링·DB 축적", "detail": "다양한 소스로 데이터 커버리지 확대"},
         {"header": "관리 (Refining)", "desc": "직급·연차 정제로 고객사 비교 가능한 형태로 가공",
          "ai": "4단계↔2단계 직급 정합", "detail": "고객사 맞춤 비교 데이터 정제"},
         {"header": "활용 (Analytics)", "desc": "직군/직무별 상대 보상 수준·%tile 분석",
          "ai": "빅데이터 기반 수준 추정", "detail": "Domain→Analyze→Implication 흐름"}],
        bar_label="데이터 확보 → 관리 → 활용")
    mark("Why HCG 5-데이터 3col")

    # ════ Ⅱ. 수행 방안 ════
    build_section_divider(prs, "Ⅱ", "수행 방안",
                          "현황진단 중심 — 벤치마킹 / 글로벌 리서치 / 구성원 의견 / 분석·Report")
    mark("Section Ⅱ")

    build_process_roadmap(prs, "Process Overview",
        "진단에 포커스한 4개 모듈을 12주에 걸쳐 단계적·병행 추진 (방향성 설계는 후속 프로젝트)",
        [{"name": "1. 벤치마킹", "steps": ["Benchmarking\nConnect", "Global Big\nData Analytics", "Local\nCompensation"]},
         {"name": "2. 글로벌 리서치", "steps": ["Regional\n노동환경", "임금체계\n전환 맥락", "법·정책\n변화"]},
         {"name": "3. 구성원 의견", "steps": ["심층 설문", "직군 인터뷰", "정서·만족도"]},
         {"name": "4. 분석·Report", "steps": ["직군별 진단", "핵심인재\n관리", "Gap·시사점\nReport"]}])
    mark("Process Overview 로드맵")

    build_body_process(prs, "현황 진단  Framework — HCG HR Maturity",
        "보상체계 성숙도(Maturity)를 기준으로 現 수준을 진단하고 개선 당위성을 단계로 제시",
        [("Lv1 연공", "호봉·연공\n획일 보상"),
         ("Lv2 직무", "직무가치\n반영 시작"),
         ("Lv3 성과", "성과 연동\n차등 보상"),
         ("Lv4 전략", "직군별 최적화\n전략적 보상")],
        [("진단 관점", "기아 現 수준을 성숙도 모델 상에 위치시키고, 직군별 격차와 개선 여지를 가시화"),
         ("활용", "개선의 '방향'이 아닌 '필요성·당위성'을 단계적 근거로 제시")])
    mark("HR Maturity Framework")

    build_body_single(prs, "수행방안 1. 벤치마킹  Overall Approach",
        "벤치마킹은 3개 모듈 — 직접 심층 BM(Connect), 글로벌 빅데이터, 국내 보상 데이터로 다층 확보",
        [("1) Benchmarking Connect", "직접 미팅·밋업·인터뷰로 제도·이슈·변화 맥락까지 심층 확인 (관행적 간접 BM과 차별)"),
         ("2) Global Big Data Analytics", "글로벌 보상 플랫폼·크롤링으로 직군/직무별 상대 수준·%tile 빅데이터 분석"),
         ("3) Local Compensation", "국내 대형 보상 플랫폼 데이터로 국내 대기업 임금 수준 비교"),
         ("통합", "정량(데이터) + 정성(맥락)을 결합해 '해석 가능한' BM 결과 도출")])
    mark("BM Overall Approach")

    build_approach_vs(prs, "수행방안 1. 벤치마킹  1) Benchmarking Connect",
        "보상 데이터를 '받아 전달'하는 관행적 BM이 아닌, 맥락을 '직접 해석'하는 HCG의 명명된 BM 방법론",
        "관행적(양산형) 벤치마킹",
        ["데이터 서베이 전담팀 활용 간접 BM", "Broad 정책·평균 %tile 중심",
         "세부 제도·변화·이슈 파악 한계", "수치(numerical) 데이터에 집중"],
        "Benchmarking Connect (직접·Context)",
        ["BM 미팅·밋업·화상으로 직접 심층 확인", "제도·변화·이슈까지 종합 파악",
         "데이터 현재성·맥락(context) 확보", "직군별 차별화 보상 포인트 해석"],
        bottom_quote="데이터를 '받아 전달'하는 BM이 아니라, 맥락을 '직접 연결(Connect)'해 시사점을 도출하는 BM")
    mark("Benchmarking Connect (VS)")

    s = build_body_single(prs, "[참고] HCG's Benchmarking Manual",
        "직접 BM의 신뢰성을 담보하는 표준 프로세스 — 질문지 설계부터 검증까지",
        [("Step 1 설계", "BM 목적·직군별 Key Question에 맞춘 질문지 설계"),
         ("Step 2 Connect", "대상사 담당자 섭외·미팅/밋업으로 직접 확인"),
         ("Step 3 검증", "수집 데이터의 현재성·정합성 교차 검증"),
         ("Step 4 해석", "맥락 기반 해석으로 기아 보상 시사점 도출")])
    add_example_badge(s)
    mark("[참고] BM Manual")

    build_compare_table(prs, "수행방안 1. 벤치마킹  BM 대상 및 범위",
        "RFP 지정 대상을 중심으로, 핵심인재 보상제도와 현장직(직고용) 구조를 포함해 조사",
        ["구분", "대상 기업 (RFP 지정)", "조사 초점"],
        [["글로벌 완성차", "도요타 · 폭스바겐 · 테슬라 · BYD", "생산직(직고용) 임금구조·수준"],
         ["국내 기업", "삼성전자 · LG전자 · SK하이닉스 · 네이버 · 카카오", "사무직 임금·핵심인재 보상"],
         ["핵심인재", "각 사 핵심인재 보상제도 (사무직군)", "영입·리텐션 보상 패키지"],
         ["협의 사항", "추가 조사 / 제외 대상", "기아와 협의 후 확정"]],
        example=True)
    mark("BM 대상 (table)")

    build_overview_3col(prs, "수행방안 1. 벤치마킹  2) Global Big Data Analytics",
        "글로벌 보상 데이터를 단순 수집이 아닌 확보→분석→시사점의 흐름으로 가공",
        [{"header": "Domain Knowledge", "desc": "글로벌 보상 플랫폼 + BM 네트워크로 데이터 확보",
          "ai": "데이터 크롤링·DB 축적", "detail": "직급·연차 정제로 비교 가능화"},
         {"header": "Analyze", "desc": "직군/직무별 상대 보상 수준·%tile 분석",
          "ai": "빅데이터 기반 수준 추정", "detail": "구조적 임금 동인 다층 분석"},
         {"header": "Implication", "desc": "기아 보상 리스크·시사점 도출",
          "ai": "맥락 기반 해석", "detail": "진단·리스크·후속 설계 활용"}],
        bar_label="Domain Knowledge → Analyze → Implication", example=True)
    mark("Global Big Data 3col")

    build_body_single(prs, "수행방안 1. 벤치마킹  3) Local Compensation",
        "국내 대형 보상 플랫폼 데이터로 국내 대기업 임금 수준을 직군/직급별로 비교",
        [("데이터 소스", "국내 대형 보상 플랫폼 데이터 (소스 확대를 위한 제휴 포함)"),
         ("정제", "기아 직급체계에 맞춘 직급·연차 비교 데이터 정제"),
         ("분석", "삼성전자·LG전자·SK하이닉스·네이버·카카오 등 국내 대기업 수준 비교"),
         ("산출", "직군/직급별 상대 보상 수준 %tile 및 격차 분석")])
    mark("Local Compensation")

    build_body_single(prs, "수행방안 2. 글로벌 리서치  Overall Approach",
        "글로벌 보상 데이터는 '받는 것'만으로 부족 — 노동환경·제도 맥락을 해석해야 시사점이 생김",
        [("1) Regional 노동환경", "각국(미·일·독) 노동법·노동환경·노조 구조 등 보상이 이뤄지는 근저 환경 리서치"),
         ("2) 임금체계 전환 맥락", "호봉제 대안으로서 직능급·성과제 전환 경로(일본), COLA(미국) 등 맥락"),
         ("3) 법·정책 변화", "통상임금·평균임금 등 국내 법·정책 변화가 임금구조에 미치는 영향"),
         ("목적", "데이터 해석을 위한 딥 리서치 — 진단·리스크 분석의 해석 기반")])
    add_insight_quote(s, "글로벌 데이터를 '제대로 해석'할 수 있는 역량이 BM의 진짜 차별점")
    mark("리서치 Overall")

    build_overview_3col(prs, "수행방안 2. 글로벌 리서치  1) Regional 노동 환경",
        "주요 완성차 보유국의 노동환경·노동법·임금 동인을 이해해야 보상구조 차이를 해석 가능",
        [{"header": "미국 (US)", "desc": "직고용·완성차 노조 단체협약, 물가연동 COLA",
          "ai": "테슬라·BYD 신흥 보상", "detail": "'20년 이후 COLA pause→부활 핫이슈"},
         {"header": "일본 (JP)", "desc": "춘투(春鬪) 베이스업, 직능급→성과제 전환",
          "ai": "도요타 직능·성과 기반", "detail": "특근수당 구조 상이(특근 60%)"},
         {"header": "독일 (DE)", "desc": "산별노조(IG Metall) 단체협약 임금",
          "ai": "폭스바겐 단협 구조", "detail": "직무·등급 기반 임금"}],
        bar_label="Regional Environment", example=True)
    mark("Regional 3col")

    s = build_body_single(prs, "수행방안 2. 글로벌 리서치  2) 임금체계 전환의 맥락  호봉제",
        "한국형 수당 중심 호봉제의 대안을 글로벌 전환 경로에서 탐색 (메인 리서치 콘텐츠)",
        [("한국", "기본급 + 연장/특근수당 + 상여 구조에서 수당 비중이 가장 높음 — 야근 50%·특근 100%"),
         ("일본", "직능급 등장 이후 성과제로 단계적 이행, 특근 60% — 전환 경로 참고"),
         ("미국", "기본급 중심 + COLA(물가연동), '20년 이후 pause 후 부활이 최근 이슈"),
         ("시사점", "기아 통상임금 확대 시 수당 더블업 리스크 → 기본급 중심 구조로의 재편 화두")])
    add_insight_quote(s, "수당 중심 호봉제의 구조적 리스크 — 글로벌 전환 경로에서 재편의 단서를 찾는다")
    mark("호봉제 맥락")

    build_compare_table(prs, "수행방안 2. 글로벌 리서치  보상구조 수평비교",
        "한·미·일 생산직 보상구조를 수평 비교하고 '왜 다른가'의 맥락으로 한국 보상 리스크를 도출",
        ["국가", "기본급", "수당/변동", "임금 결정 메커니즘", "시사점"],
        [["한국", "비중 中", "연장/특근수당 高·상여", "총액 기준·호봉 상승", "통상임금 더블업 리스크"],
         ["미국", "비중 高", "COLA(물가연동)", "단협·COLA pause→부활", "기본급 중심 전환 참고"],
         ["일본", "직능 기반", "특근 60%", "춘투 베이스업·성과제 이행", "단계적 전환 경로"]],
        example=True)
    mark("보상구조 수평비교 (table)")

    s = build_body_single(prs, "수행방안 2. 글로벌 리서치  3) 법·제도 및 정책 변화",
        "국내 법·정책 변화가 임금구조에 미치는 영향을 리서치해 리스크 진단의 근거로 활용",
        [("통상임금", "산입 범위 확대 판례·정책 변화 추적"),
         ("평균임금", "퇴직급여·충당부채 영향 분석"),
         ("근로시간", "연장·야간·휴일 근로 및 수당 관련 제도 변화"),
         ("활용", "법·정책 Risk를 정량화해 보상 재편의 당위성 근거로 연결")])
    add_example_badge(s)
    mark("법·정책 변화")

    build_body_single(prs, "수행방안 3. 구성원 의견  심층 설문 · 인터뷰",
        "정량 BM에 정성 조사를 결합해 진단 신뢰성을 제고 (조사 대상 한정)",
        [("대상", "사무직군 중심 — 생산직 인터뷰는 아님 (RFP상 대상 한정 명확화)"),
         ("심층 설문", "보상 만족도·직원 정서에 대한 심층 설문 진행"),
         ("인터뷰", "직군별 담당자 인터뷰로 제도 현황·이슈·수용성 파악"),
         ("활용", "정량(BM) + 정성(의견) 결합으로 진단 결과의 설득력 강화")])
    mark("구성원 의견")

    build_diff_matrix(prs, "수행방안  분석 및 Report  직군별 보상 진단",
        "직군별 업무·보상 특성에 따라 진단의 Key Question을 차별화 (가설적 구분)",
        [{"group": "일반/연구직\n(사무)", "trait": "국내 대기업 임금경쟁, 핵심인재 리텐션",
          "insight": "시장 경쟁력·리텐션 관점", "apply": "국내 대기업 BM + 핵심인재 패키지 진단"},
         {"group": "엔지니어\n(생산)", "trait": "호봉·수당 중심, 통상임금 리스크",
          "insight": "구조적 임금 동인·호봉 리스크 관점", "apply": "글로벌 완성차 직고용 BM, 호봉제 실태"},
         {"group": "기술직\n(정비)", "trait": "서비스센터 현장 정비, 직무 특수성",
          "insight": "직무 특성 기반 적정성 관점", "apply": "유사 직무 BM, 직무가치 진단"},
         {"group": "오토컨설턴트\n(영업)", "trait": "지점 영업, 성과 변동성",
          "insight": "성과 연동 적정성 관점", "apply": "영업 직군 성과보상 BM"}],
        bottom_quote="직군별로 보상 Key Question이 달라야 한다 — 일반적 4P가 아닌 직군별 차별화 진단",
        example=True)
    mark("직군별 진단 matrix")

    s = build_body_single(prs, "[참고] 기아 매니저급 임금체계 변화",
        "노조 가입 여부에 따른 보상구조 차이가 사무직 매니저급의 형평성 이슈로 부각",
        [("책임급", "노조 자동 탈퇴 → 별도 보상구조, 성과평가 적용"),
         ("매니저급", "노조원(생산직과 동일 보상구조) → 성과평가 미적용"),
         ("이슈", "사무직 매니저급의 생산직 동일 보상구조에 대한 형평성 불만"),
         ("진단 포인트", "통상임금 확대로 수당 더블업 리스크 — 매니저급 임금체계 재편 필요성")])
    add_example_badge(s)
    mark("[참고] 매니저 임금")

    s = build_body_single(prs, "수행방안  분석 및 Report  핵심인재 관리 현황 진단",
        "시장에서 희소한 고직능 인재의 영입·리텐션 경쟁력과 보상 적정성을 진단",
        [("시장 속성", "시장 희소·고직능 매니저급의 영입/리텐션 경쟁 심화"),
         ("진단", "핵심인재 영입·리텐션 경쟁력 및 보상 적정성 진단"),
         ("BM", "글로벌·국내 핵심인재 보상 패키지 조사 (사무직군 대상)"),
         ("시사점", "별도 관리·차별적 보상 패키지 필요성 도출 (후속 설계 연결)")])
    mark("핵심인재 진단")

    s = build_body_single(prs, "[참고] 핵심인재 보상 사례  테슬라",
        "파격적 보상 패키지로 핵심인재를 확보하는 글로벌 사례 — 기아 관점의 해석",
        [("보상 패키지", "기본급 + 주식보상(RSU/스톡옵션) 등 장기 인센티브 중심 구조"),
         ("리텐션", "장기 베스팅으로 핵심인재 리텐션 강화"),
         ("시장 속성", "고직능·희소 인재에 대한 시장 경쟁 보상"),
         ("해석", "기아 핵심인재(매니저급) 보상 패키지 설계 시 참고할 시장 기준")])
    add_example_badge(s)
    mark("[참고] 테슬라")

    build_body_process(prs, "수행방안  분석 및 Report  Gap 분석",
        "기아 현황과 BM·리서치 결과를 비교해 구조적 문제와 개선 필요사항을 도출",
        [("현황 분석", "기아 보상\n구조 상세"),
         ("BM 비교", "타사 임금\n구조/수준"),
         ("Gap 도출", "구조적 문제\n개선 필요사항"),
         ("시사점", "리스크·개선\n당위성")])
    mark("Gap 분석")

    build_body_single(prs, "수행방안  분석 및 Report  진단 Report 구성",
        "진단 결과를 활용 가능한 Report로 산출 — 개선 방향성 설계는 후속 프로젝트",
        [("① 벤치마킹 결과 Report", "대상별·직군/직무별 보상 수준·구조 비교, %tile 페이밴드"),
         ("② 구성원 의견 종합 Report", "보상 만족도·정서 및 직군별 인터뷰 시사점"),
         ("③ 진단 결과 종합 Report", "제도 Gap·구조적 문제·법 리스크 종합 진단"),
         ("결론", "개선 필요성·당위성 화두 제시 (방향성·설계는 후속)")])
    mark("진단 Report 구성")

    s = build_body_single(prs, "[참고] 진단 Report Output Image",
        "RFP 요청에 따른 산출물 이미지 — 직군/직급별 보상 분포 및 구조 비교",
        [("페이밴드 Report", "직군/직급별 보상 분포·%tile 밴드"),
         ("연봉 평가 Report", "시장 대비 보상 경쟁력 평가"),
         ("구조 비교 Report", "한·미·일 보상구조 수평 비교 미니리포트"),
         ("핵심인재 Report", "핵심인재 보상 패키지 비교")])
    add_example_badge(s)
    mark("[참고] Report Output")

    # ════ Ⅲ. 일정·조직·비용 ════
    build_section_divider(prs, "Ⅲ", "추진 일정 · 조직 · 비용",
                          "12주 추진 일정, 자동차·C&B 전문 인력, 옵션형 가격 제안")
    mark("Section Ⅲ")

    build_body_process(prs, "추진 일정 (12주)",
        "9월말 Kick-off ~ 연말 경영층 보고 — 중간보고·결과보고 운영",
        [("W1-2\nKO·자료", "현황자료\n수집·분석"),
         ("W3-6\nBM·리서치", "글로벌/국내\nBM·리서치"),
         ("W7-9\n진단", "직군별 진단\nGap 분석"),
         ("W10-11\nReport", "Report 작성\n시사점 도출"),
         ("W12\n보고", "결과보고\n경영층 보고")],
        [("보고 체계", "중간보고(진단 진행) · 결과보고(경영층) 2회 운영"),
         ("산출물", "벤치마킹 결과·구성원 의견·진단 종합 Report")])
    mark("추진 일정")

    build_body_2col(prs, "수행 조직",
        "기아 HR지원팀과 HCG 전문 컨설턴트의 긴밀한 협업 체계 (핵심 3명 / 가용 6명)",
        "기아 HR지원팀", "HCG 수행팀",
        ["프로젝트 스폰서 / 의사결정", "BM 대상·자료 협조",
         "구성원 조사 협조", "중간·결과보고 검토"],
        ["PM — 보상 컨설팅 총괄", "BM 리드 — 글로벌/국내",
         "데이터 분석 — 크롤링·빅데이터", "직군별 진단 컨설턴트"])
    mark("수행 조직")

    build_body_single(prs, "투입 인력  PM · Member",
        "자동차 산업·대기업 C&B 경험 전문가 중심 (실제 자동차/보상 프로젝트 경험 기재)",
        [("PM", "대기업 보상 컨설팅 총괄 — 현대차·기아 등 자동차 보상 PJT 수행"),
         ("Member 1 (BM)", "글로벌/국내 벤치마킹 리드 — 파라다이스·SK하이닉스 BM 경험"),
         ("Member 2 (분석)", "보상 데이터 크롤링·빅데이터 분석"),
         ("가용 인력", "TH·CY·CG 등 가용 6명 — 자동차·C&B 전문가 풀 보유")])
    mark("투입 인력")

    build_body_single(prs, "가격 제안",
        "프로젝트 수행 경비 포함, 비용 지급시점 협의 — 옵션형으로 제안",
        [("옵션 1 (진단)", "본 프로젝트 범위 — 12주, 2억 초반대 (경비 포함·부가세 별도)"),
         ("옵션 2 (후속 포함)", "후속 제도개선·방향성 설계까지 패키지 가견적 (협의)"),
         ("BM 방식", "BM 미팅·밋업·화상 포함 — 직접 방문(Trip)은 협의 옵션 (트립 디폴트 아님)"),
         ("지급 시점", "착수·중간·완료 등 비용 지급 시점 협의")])
    mark("가격 제안")

    # ════ Ⅳ. HCG 소개 (간략) ════
    build_section_divider(prs, "Ⅳ", "HCG 소개", "국내 최대·유일의 HR End-to-End 전문가 그룹")
    mark("Section Ⅳ")

    build_body_single(prs, "HCG 소개",
        "보상(C&B) 전문성과 자동차 산업 경험을 모두 보유한 HR 컨설팅 그룹",
        [("정체성", "국내 최대·유일의 HR End-to-End 전문가 그룹"),
         ("보상 전문성", "직군별 차별화 C&B 설계·벤치마킹 다수 수행"),
         ("자동차 산업", "현대차·기아·현대오토에버 등 완성차 생태계 보상/성과 경험"),
         ("주요 고객사", "기아·현대차·두산인프라코어·코오롱·GS·녹십자 등 870여 곳")])
    mark("HCG 소개")

    build_end(prs); mark("End")

    print(f"\nSaving → {OUT}")
    prs.save(OUT)
    print("Done! 총", n[0], "장")
    return OUT


if __name__ == "__main__":
    main()
