# -*- coding: utf-8 -*-
"""
HCG 제안서 자동 PPT 생성기 v2.0
RFP: 롯데알미늄 직무기반 HR제도 설계 및 도입 컨설팅
생성 결과: HCG_Automated_Draft.pptx

변경 이력 v2.0:
  - 슬라이드 크기: 13.33" → 10.83" × 7.50" (한국 표준)
  - 본문 텍스트: bold blue → Regular black (#000000)
  - 구조: _final 25장 구조 반영 (레퍼런스 앞배치, Project Overview 통합 등)
  - [참고] 슬라이드: 헤드 없음, 타이틀만

실행: python auto_ppt.py
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

# ─────────────────────────────────────────────
# 색상 상수
# ─────────────────────────────────────────────
DEEP_BLUE    = RGBColor(0x00, 0x38, 0x87)   # 헤드 메시지, 번호, 구분선
LIGHT_BLUE   = RGBColor(0x5B, 0x9B, 0xD5)   # 서브 강조, 헤더 바
BODY_BLACK   = RGBColor(0x1A, 0x1A, 0x1A)   # 본문 텍스트 (≈ #000000)
GRAY         = RGBColor(0x80, 0x80, 0x80)    # 섹션 라벨, 메타
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)    # 커버 텍스트
HEAD_BG      = RGBColor(0xF2, 0xF6, 0xFF)   # 헤드 배경

# ─────────────────────────────────────────────
# 슬라이드 데이터 (_final 25장 구조 반영)
# ─────────────────────────────────────────────
SLIDES = [
    {
        "no": 1,
        "type": "cover",
        "section": "Cover",
        "head": "롯데알미늄 직무기반 HR 제도 설계 및 도입",
        "body": [
            "- 제안서 -",
            "",
            "2026.03",
        ],
    },
    {
        "no": 2,
        "type": "toc",
        "section": "목차",
        "head": "Contents",
        "body": [
            "I.    제안 개요 및 배경",
            "II.   수행 방안 (직무체계 / 평가 / 보상)",
            "III.  추진 계획",
            "Appendix.  HCG 소개",
        ],
    },
    {
        "no": 3,
        "type": "normal",
        "section": "I. Project Overview",
        "head": "Project Overview",
        "sub": "롯데알미늄의 직무체계 및 평가/보상제도 운영방안 도출",
        "body": [
            "[추진 내용] 직무체계 및 평가/보상제도 운영방안 도출",
            "[추진 프로세스] 현황 진단 및 방향성 수립 → 직무분류체계 수립 및 직무기술서 도출 → 보상/평가제도 개선",
            "[추진 방안(잠정)] ① 그룹사 직무 Guideline 현황 Quick Review·개선방향 도출",
            "② AI 활용 직무분류체계 초안 도출 및 Fine-tuning/Customization·직무기술서 작성",
            "③ 직무/업무 특성을 반영한 성과관리체계 수립·평가 운영체계 개선",
            "④ 보상 원칙/전략 및 보상 운영방안 수립·보상 재원 Simulation",
            "",
            "[추진 기간] 12주  |  [투입 인력] PM 1명 + 컨설턴트 2명  |  [추진 금액] 225,000천원(VAT별도)",
        ],
    },
    {
        "no": 4,
        "type": "normal",
        "section": "I. Client's Needs & Pain Point",
        "head": "Client's Needs & Pain Point",
        "sub": "직무체계/평가 및 보상제도 미흡에 따른 인사제도 개편이 필요하나, 기존의 컨설팅 방식은 많은 자원이 장기간 투입되어 효율적인 HR제도 정비가 어려움",
        "body": [
            "[파편화된 데이터와 수동적 HR 운영의 악순환]",
            "과거 History·현황 파악 등에 많은 시간/비용 소요",
            "설계 과정에서 Interview·Survey 등 구성원에게 부담",
            "일부 직군·구성원은 수용성이 떨어져 평가 결과에 불만",
            "인력별 업무/직무 가치 파악 어려움 / 현행화되지 않은 직무 정보",
            "",
            "[유기적 연결이 부족한 컨설팅 결과물]",
            "객관적 데이터 부재 : 기준이 없거나 노후화된 데이터로 HR 제도에 대한 구성원의 신뢰/수용성 하락",
            "High Cost, Low Return : 막대한 리소스 투입 대비 실무 적용이 어려운 산출물",
            "사후관리 부재 : 변화관리가 이뤄지지 않아 기존 방식으로 회귀",
        ],
    },
    {
        "no": 5,
        "type": "normal",
        "section": "I. HCG's Approach",
        "head": "HCG's Approach",
        "sub": "HR 전문가의 지식/경험과 글로벌 표준을 집대성한 '표준 설계도'를 기반으로 AI를 활용해 직무 기반 HR제도 개선에 대한 Pain Point를 쉽고, 정확하고, 효율적인 방법으로 해소함",
        "body": [
            "[AI를 활용해 즉시 활용 가능한 고객 맞춤형 직무체계 구축 및 직무 특성/가치를 반영한 평가/보상제도 수립 가능]",
            "",
            "기존 컨설팅 방식: 수작업 기반 / 시간·비용 과다 / 현행화 한계 / 사후관리 부재",
            "VS",
            "HCG AI Approach: 표준을 빠르게 만들고 → 고객과 Customize → 납득성 높은 결과 도출",
            "핵심: '텍스트가 아니라 표준을 빠르게 만든다' — AI가 초안, 전문가와 협업해 Customization",
        ],
    },
    {
        "no": 6,
        "type": "ref",
        "section": "I. [참고]",
        "head": "[참고] 高맥락 컨설팅 접근 방식",
        "body": [
            "HCG는 겉으로 드러난 문맥만으로는 반영하기 어려운 고객사 고유의 비언어적 맥락(HR Gene)을 이해/반영하고, 이를 HR 모듈에 즉각 적용함",
            "",
            "[해소 방안 ①] 회사가 가지고 있는 암묵적 'HR Gene'을 이해하고 Context로서 활용",
            "  — 그룹의 인재 활용 철학·조직문화·구성원 경험의 변화사 등 맥락에 대한 이해 선행 필요",
            "",
            "[해소 방안 ②] 필요에 따라 다양한 비정형적 Data 활용",
            "  — HR 모듈별 필요에 따라 비정형 Data를 유연하게 활용하여 고객사 최적화 HR제도 운영",
        ],
    },
    {
        "no": 7,
        "type": "normal",
        "section": "I. 유사 프로젝트 수행 사례",
        "head": "유사 프로젝트 수행 사례",
        "sub": "국내 대기업 및 다수의 기업 대상 직무 기반의 HR제도 컨설팅 수행 실적을 통해 관련 주제에 대한 폭넓은 컨설팅 경험을 토대로 최적의 프로젝트 결과를 제공할 것임",
        "body": [
            "직무기반 조직재설계 / 조직 재설계(직무 설계 포함)",
            "조직구조 재설계 및 직무별 인력운영 구축",
            "직무 기반 보상·평가 및 인사제도 전반 개선",
            "직군별 인사관리 및 인력 효율화",
            "직무성과주의 평가/보상제도 구축",
            "직무분석 및 인사제도 개선",
            "조직/직무 설계 및 인력계획 수립",
            "조직 및 직무·평가·보상제도 개선",
            "직무 및 업무분석을 통한 인력 효율화",
            "직무체계 및 평가제도 개선",
        ],
    },
    {
        "no": 8,
        "type": "normal",
        "section": "II. 직무체계 개선    Overview",
        "head": "직무체계 개선    Overview",
        "sub": "지주사의 직무 분류 철학을 계승하되, AI 기반의 글로벌 벤치마킹과 생산 현장의 특수성을 반영하여 신속하게 롯데알미늄에 최적화된 표준 직무체계를 수립할 것을 제안함",
        "body": [
            "Step 1. Analysis — 직무 조사 및 분석",
            "  구성원 대상 수행 중인 업무 목록·소요 시간·필요 역량 등 작성 요청",
            "  SME 인터뷰를 통해 검증/조정",
            "",
            "Step 2. Classification — 직군 정의 및 직무 도출",
            "  산업의 가치사슬·업무의 성격 등을 기준으로 유사한 대분류 설정",
            "  직무 정의·목표·과업·자격 요건 등 확인",
            "",
            "Step 3. Evaluation — 직무가치 평가/등급 도출",
            "  직무 가치를 기준에 따라 점수화해 평가",
            "  도출된 점수를 바탕으로 직무 등급화(평가/보상 연계 가능)",
        ],
    },
    {
        "no": 9,
        "type": "normal",
        "section": "II. 직무체계 개선    Process",
        "head": "직무체계 개선    Process",
        "sub": "'AI + HCG DB'를 통해 직무를 빠르게 분석 후 Communication Lead Time을 최소화해 고객 요구사항이 내재화된 직무체계를 수립하고, 이를 기반으로 직무평가를 진행함",
        "body": [
            "[지주사 Guideline] 지주사 직무분류 원칙을 롯데알미늄 비즈니스 맥락에 맞게 해석/적용",
            "[AI-base Benchmark] AI를 활용해 종합 포장 소재 산업의 Global Standard 및 Best Practice 수집",
            "[JobStandardSetting] 직무분류체계 + 직무기술서 → Agile Delivery",
            "[Context & History] 현장 맥락·과거 이력 반영 → Finalize",
            "",
            "산출물: 직무분류체계(XLSX) / 직무기술서(XLSX)",
        ],
    },
    {
        "no": 10,
        "type": "normal",
        "section": "II. 직무체계 개선    직무체계 표준화",
        "head": "직무체계 개선    직무체계 표준화",
        "sub": "공장별로 상이했던 직무 기준을 Global Standard에 맞춰 통합함으로써, 직무 운영의 비효율을 해소하고 일관된 평가/인재관리 효과성을 제고할 것임",
        "body": [
            "[파편화된 직무체계 및 표준화 한계]",
            "공장별 독자적 업무 수행으로 직무분류체계 표준화 어려움(직무 단위 불일치)",
            "직무분류 기준 부재로 직무 중복·비효율적 인력 배치 발생",
            "공통된 요구 역량 기준 없어 일관된 평가/육성 연계 한계",
            "",
            "[Global Standard 기반 직무/Skill 표준화 — 개선 방향]",
            "Global 표준 직무 정의/Skill Set 등 기준 명확화",
            "각 공장의 Value Chain과 업무 방식을 新 표준에 맞게 재편(비효율 및 중복 해소)",
            "직무별 요구되는 Skill 정의 및 수준을 전사 공통으로 규격화",
        ],
    },
    {
        "no": 11,
        "type": "ref",
        "section": "II. [참고]",
        "head": "[참고] Global Job Skill 분석 결과 예시",
        "body": [
            "HCG JobSkillEngine (Claude API 기반) 활용 시 산업별 직무분류체계·직무기술서·스킬디렉토리 자동 생성 결과 예시",
            "",
            "분석 기준: APQC / ESCO / ISCO / O*NET / NCS 5개 국제 표준 참조",
            "직군→직렬→직무 3단계 분류체계 샘플",
            "스킬 디렉토리 샘플 (직무별 KSAO, 스킬 클러스터, PMI 공동출현 분석 포함)",
        ],
    },
    {
        "no": 12,
        "type": "normal",
        "section": "II. 평가제도 개선    Overview",
        "head": "평가제도 개선    Overview",
        "sub": "평가제도 설계 시 도출된 직무체계를 바탕으로 직무에 특화된 평가 모델을 구축해 운영 프로세스를 개선할 수 있는 방향성을 검토함",
        "body": [
            "Step 1. Criteria — 평가 원칙 수립",
            "  평가 종류(성과/역량)·목표 관리 프레임워크(KPI)·직급/직군별 가중치 등 평가 지표 및 체계에 대한 원칙 논의",
            "",
            "Step 2. Process — 평가 운영체계 개선",
            "  평가 권한·단위·등급·프로세스·모니터링 등 평가 운영 전반에 대한 프로세스 개편",
            "  필요 시 성과관리 시스템 기반 제도 구축",
            "",
            "Step 3. Application — 결과 활용 및 타 Module 연계 방안",
            "  보상 차등의 기준으로 활용 가능",
            "  필요 시 리더십을 포함한 인재 육성 및 저성과자 관리 방안 제시",
        ],
    },
    {
        "no": 13,
        "type": "normal",
        "section": "II. 평가제도 개선    직군별 차별화 평가모델",
        "head": "평가제도 개선    직군별 차별화 평가모델",
        "sub": "\"각 직군별 업무 특성을 반영한 평가제도 마련으로 인력 관리의 효과성 제고 및 성과 향상 지원\"",
        "body": [
            "[R&D 및 신사업] 프로젝트 마일스톤 Output 중심 평가·R&D 장기 성과급 체계 도입",
            "  → 마일스톤 중심의 장기 성과 인정, 기술 경쟁력 확보 관점의 인력 관리",
            "",
            "[영업] 채널 난이도에 따른 가중치 반영 평가·개인 성과 비율 확대(인센티브 비중 강화)",
            "  → 목표 달성에 따른 성과/보상연계 강조 및 개인 성과 인정",
            "",
            "[생산] 절대적 달성 수준 기반 평가(절대평가 고려)·조직 성과와 연계된 보상제도 수립",
            "  → 조직 단위 성과 발현, 반복/절대적 수행 업무 관리 중요",
            "",
            "[경영지원] 핵심 과제 중심으로 목표 설정·팀간 협업 지표를 목표에 반영",
            "  → 핵심 과제 중심의 상시성과관리 및 협업에 대한 정성 목표 설정 체계화",
        ],
    },
    {
        "no": 14,
        "type": "ref",
        "section": "II. [참고]",
        "head": "[참고] AI 기반 업무 Check-in 시스템 활용 예시 (사무직 한정)",
        "body": [
            "HCG 자체 상시성과관리 시스템 talenx — AI 기반 목표 설정/체크인/피드백 지원",
            "",
            "직원별 직무/목표/성과 Data를 분석해 조직 성과 방향과 정렬된 핵심성과지표 제안",
            "AI를 활용한 목표 설정 / 체크인 / 피드백까지 데이터 기반의 평가 체계 전환에 용이",
            "",
            "※ 체크인(상시성과관리)은 사무직에 한하여 적용 — 생산직은 KPI 기반 절대평가 적용",
        ],
    },
    {
        "no": 15,
        "type": "normal",
        "section": "II. 평가제도 개선    평가운영체계",
        "head": "평가제도 개선    평가운영체계",
        "sub": "정비된 평가제도에 따라 평가 등급, 평가 프로세스, 평가 결과 활용 등의 운영체계를 개선함",
        "body": [
            "평가 등급 기준 재정립 (절대/상대 혼합 기준 설계)",
            "평가 프로세스 단계별 가이드: 목표설정 → 중간점검 → 최종평가 → 결과 통보",
            "이의신청 온라인 채널 구축",
            "평가 결과의 보상 연계 비율 및 기준 명확화",
            "저성과자 관리 프로세스 수립",
        ],
    },
    {
        "no": 16,
        "type": "normal",
        "section": "II. 보상제도 개선    Overview",
        "head": "보상제도 개선    Overview",
        "sub": "보상제도 설계 시 보상의 외부 경쟁력 및 내부 형평성을 종합적으로 고려하여 보상 수준, 보상 구조, 보상 결정요인, 성과급 비중 등 개선 방향성을 검토함",
        "body": [
            "Step 1. Pay Level — 보상 수준 조사 및 보상 전략 도출",
            "  경쟁업체와 비교 시 인재 확보/유지를 위한 보상 경쟁력 확인",
            "  직무/역할 특성을 고려한 내부 형평성 고려",
            "",
            "Step 2. Pay Structure / Mix / Contributor — 보상 구조/비중/결정요인 체계화",
            "  전체 보상을 구성하는 항목/수준 검토(기본급·전사/개인 성과급 등)",
            "  보상을 결정 짓는 요소 선정",
            "",
            "Step 3. Simulation — Cost Impact / 개인 총 보상 검증",
            "  개편될 보상제도를 기준으로 단기/중장기 비용 측면에서 Simulation",
            "  구성원 개인별 총 보상 변화 양상 검토",
        ],
    },
    {
        "no": 17,
        "type": "normal",
        "section": "II. 보상제도 개선    보상 지향점/정책선 설정",
        "head": "보상제도 개선    보상 지향점/정책선 설정",
        "sub": "국내외 동종업계의 보상수준 Benchmarking을 통해 보상 경쟁력을 검토하고 보상 지향점/정책선을 설정함",
        "body": [
            "Market Leading Line (75%ile) — 시장 선도 보상 수준",
            "Market Matching Line (50%ile) — 시장 동등 보상 수준",
            "Market Following Line (25%ile) — 시장 후발 보상 수준",
            "",
            "동종업계의 보상수준 Benchmarking 통해 보상수준 비교 및 지향점 분석 시 활용 예정",
        ],
    },
    {
        "no": 18,
        "type": "normal",
        "section": "II. 보상제도 개선    보상제도 설계",
        "head": "보상제도 개선    보상제도 설계",
        "sub": "보상제도 설계 시 보상 비중/유형/지급 기준 등을 종합적으로 고려하며, 필요 시 직급단계 혹은 시장가치를 고려하여 보상 차별화 필요성을 검토할 것임",
        "body": [
            "보상 전략과 원칙을 바탕으로 보상 구조/비중/결정요인을 다각도로 분석 및 설계",
            "",
            "[직무 차별화]",
            "직무별 Market 수준에 따라 기본급 Band 차별적 운영",
            "직무 중심 차별화에 강력한 수단 / 직무 간 이동·배치에 제약",
            "",
            "[직급 차별화]",
            "직급별 내부 가치 및 개인 역량에 따른 차별화 실시",
            "동일 직급 내 개인의 성과/역량에 따른 보상 차별화 유리",
            "",
            "[직무 & 직급 혼합] 두 방식의 장점을 결합하여 롯데알미늄 맥락에 최적화된 보상 설계",
        ],
    },
    {
        "no": 19,
        "type": "normal",
        "section": "II. 보상제도 개선    보상 Simulation",
        "head": "보상제도 개선    보상 Simulation",
        "sub": "보상 재설계에 따른 개인별 보상 수준의 변화를 확인하고, 이에 따른 재무적 Cost를 Simulation하여 조직 및 구성원 개인 관점의 효과성을 검증함",
        "body": [
            "[Scenario 변동 요소]",
            "성과급 비율 / 기본급 인상률 / 직무 등급별 Pay Band 범위",
            "",
            "[전사 재무적 효과 Cost Simulation]",
            "성과급 제도 실시에 따른 전사 및 기능/사업별 Cost 증감 효과 Simulation",
            "",
            "[개인별 보상전환 Simulation]",
            "To-Be 성과급 제도 적용 시 개인별 성과급 수준 변화 Simulation",
            "",
            "[Scenario List] 보수적 / 중립적 / 공격적 3개 시나리오",
            "Illustrative use only",
        ],
    },
    {
        "no": 20,
        "type": "toc",
        "section": "Appendix 구분",
        "head": "Contents",
        "body": [
            "Appendix.  HCG 소개",
        ],
    },
    {
        "no": 21,
        "type": "normal",
        "section": "Appendix. Overview",
        "head": "Overview",
        "sub": "휴먼컨설팅그룹(이하 HCG)는 국내 최고/최대의 인사/조직 컨설팅 및 솔루션 전문 업체임",
        "body": [
            "2001년 1월 법인 설립",
            "컨설턴트 100명+ / 누적 고객사 870여 곳+ (2025년 01월 기준)",
            "HR Solution (talenx·퍼플 등 자체 플랫폼 운영)",
        ],
    },
    {
        "no": 22,
        "type": "normal",
        "section": "Appendix. 사업 영역",
        "head": "사업 영역",
        "sub": "HCG는 한국적 기업 환경 이해와 인사·조직 전문성에 기반하여 경영 컨설팅과 솔루션 컨설팅, HR System Management/Outsourcing을 통합 서비스하는 HR Total Solution Provider임",
        "body": [
            "[Mission] Make the Business Strategy Work Through HR Total Solution",
            "[Strategy] HCG는 제도 컨설팅 뿐만 아니라 e-HR Solution, PO Service 등 HR Value Chain",
            "  전 영역에 걸쳐 수준 높은 통합 서비스를 제공하여 고객의 궁극적인 비즈니스 전략 실행 및 성공을 지원",
        ],
    },
    {
        "no": 23,
        "type": "normal",
        "section": "Appendix. 컨설팅 영역",
        "head": "컨설팅 영역",
        "sub": "HCG는 핵심가치와 문화에 기반한 조직과 인재, 성과와 보상, Digital Analytics 등 HR 전 영역을 아우르는 컨설팅 서비스를 제공",
        "body": [
            "[Performance & Reward] 성과관리 혁신·역량관리·경력개발·총 보상 전략·임원 보상·글로벌 HR",
            "[Organization & Talent] 조직구조 설계·BPI·HR 거버넌스·인력계획·직무 중심 인재관리·채용 전략",
            "[Corp. Value & Culture] 기업 가치체계 수립·고용 브랜드 개선·직원 경험 관리·직원 인식 조사",
            "[HR Transformation] HR 전략 수립·HR 서비스 모델 혁신·변화관리",
            "[Digital Analytics] 디지털 HR전략·HR 테크놀로지·HR 데이터 분석·실시간 개인화 리포팅",
        ],
    },
    {
        "no": 24,
        "type": "normal",
        "section": "Appendix. 주요 고객사",
        "head": "주요 고객사",
        "sub": "HCG는 다양한 산업영역에 걸쳐 870여 곳 이상의 고객사를 대상으로 Total HR Service를 제공하고 있음",
        "body": [
            "금융: KB라이프생명 / NH투자증권 / 파라다이스 / 한화손해보험 등",
            "제조: 기아자동차 / 현대자동차 / 코오롱인더스트리 / 포스코인터내셔널 / 롯데 그룹사 등",
            "IT·게임: 크래프톤 / SK머티리얼즈 / GS에너지 등",
            "유통·소비재: 롯데백화점 / 코웨이 / 하이트진로음료 / 리파인 등",
        ],
    },
    {
        "no": 25,
        "type": "end",
        "section": "End",
        "head": "",
        "body": [],
    },
]

# ─────────────────────────────────────────────
# 헬퍼 함수
# ─────────────────────────────────────────────
def add_shape_rect(slide, left, top, width, height, fill_color, line_color=None):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape


def add_textbox(slide, text, left, top, width, height,
                font_size=11, bold=False, color=BODY_BLACK,
                align=PP_ALIGN.LEFT, wrap=True):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txBox


# ─────────────────────────────────────────────
# 슬라이드 빌더
# ─────────────────────────────────────────────
def build_cover(prs, d):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    W, H = prs.slide_width, prs.slide_height
    panel_w = Inches(4.0)

    add_shape_rect(slide, 0, 0, panel_w, H, DEEP_BLUE)
    add_shape_rect(slide, panel_w, 0, W - panel_w, H, WHITE)

    add_textbox(slide, "휴먼컨설팅그룹",
                Inches(0.3), Inches(1.0), panel_w - Inches(0.4), Inches(0.4),
                font_size=12, bold=False, color=LIGHT_BLUE)
    add_textbox(slide, d["head"],
                Inches(0.3), Inches(1.6), panel_w - Inches(0.4), Inches(1.5),
                font_size=18, bold=True, color=WHITE)

    sep = add_shape_rect(slide, Inches(0.3), Inches(3.4), Inches(1.2), Pt(2), LIGHT_BLUE)

    body_text = "\n".join([b for b in d["body"] if b])
    add_textbox(slide, body_text,
                Inches(0.3), Inches(3.6), panel_w - Inches(0.4), Inches(1.5),
                font_size=12, bold=False, color=WHITE)

    add_textbox(slide, "CONFIDENTIAL",
                panel_w + Inches(0.3), H - Inches(0.4), W - panel_w - Inches(0.4), Inches(0.3),
                font_size=8, bold=False, color=GRAY, align=PP_ALIGN.RIGHT)


def build_toc(prs, d):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    W, H = prs.slide_width, prs.slide_height

    add_shape_rect(slide, 0, 0, W, Inches(0.15), DEEP_BLUE)
    add_shape_rect(slide, 0, H - Pt(2), W, Pt(2), DEEP_BLUE)

    add_textbox(slide, d["head"],
                Inches(0.4), Inches(0.5), W - Inches(0.8), Inches(0.6),
                font_size=18, bold=True, color=DEEP_BLUE)

    add_shape_rect(slide, Inches(0.4), Inches(1.3), Pt(2), H - Inches(1.8), LIGHT_BLUE)

    for i, item in enumerate(d["body"]):
        add_textbox(slide, item,
                    Inches(0.7), Inches(1.5 + i * 0.65), W - Inches(1.0), Inches(0.55),
                    font_size=13, bold=False, color=BODY_BLACK)


def build_normal(prs, d):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    W, H = prs.slide_width, prs.slide_height

    # 상단 헤더 바
    add_shape_rect(slide, 0, 0, W, Inches(0.15), DEEP_BLUE)
    # 하단 라인
    add_shape_rect(slide, 0, H - Pt(2), W, Pt(2), DEEP_BLUE)

    # 슬라이드 번호 (좌상단)
    add_textbox(slide, f"{d['no']:02d}",
                Inches(0.15), Inches(0.18), Inches(0.5), Inches(0.28),
                font_size=9, bold=True, color=DEEP_BLUE)

    # 섹션 라벨 (우상단)
    add_textbox(slide, d["section"],
                W - Inches(3.5), Inches(0.18), Inches(3.35), Inches(0.28),
                font_size=9, bold=False, color=GRAY, align=PP_ALIGN.RIGHT)

    # 헤드 메시지
    head_top = Inches(0.55)
    head_h   = Inches(0.55)
    add_shape_rect(slide, Inches(0.3), head_top, W - Inches(0.6), head_h,
                   HEAD_BG, LIGHT_BLUE)
    add_textbox(slide, d["head"],
                Inches(0.4), head_top + Inches(0.05), W - Inches(0.8), head_h - Inches(0.08),
                font_size=13, bold=True, color=DEEP_BLUE)

    # 서브헤드 (있는 경우)
    body_top = head_top + head_h + Inches(0.1)
    if d.get("sub"):
        sep = add_shape_rect(slide, Inches(0.3), body_top, W - Inches(0.6), Pt(1.5), LIGHT_BLUE)
        add_textbox(slide, d["sub"],
                    Inches(0.35), body_top + Inches(0.05), W - Inches(0.7), Inches(0.45),
                    font_size=11, bold=False, color=BODY_BLACK)
        body_top += Inches(0.6)

    # 본문
    txBox = slide.shapes.add_textbox(Inches(0.4), body_top, W - Inches(0.7),
                                      H - body_top - Inches(0.3))
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, bullet in enumerate(d["body"]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(3)
        run = p.add_run()
        run.text = bullet
        run.font.size = Pt(11)
        run.font.bold = False         # ★ Regular (v2 수정)
        run.font.color.rgb = BODY_BLACK  # ★ 검정 (v2 수정)


def build_ref(prs, d):
    """[참고] 슬라이드 — 타이틀만, 헤드 메시지 없음"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    W, H = prs.slide_width, prs.slide_height

    add_shape_rect(slide, 0, 0, W, Inches(0.15), LIGHT_BLUE)
    add_shape_rect(slide, 0, H - Pt(2), W, Pt(2), LIGHT_BLUE)

    add_textbox(slide, f"{d['no']:02d}",
                Inches(0.15), Inches(0.18), Inches(0.5), Inches(0.28),
                font_size=9, bold=True, color=LIGHT_BLUE)

    add_textbox(slide, d["head"],
                Inches(0.4), Inches(0.5), W - Inches(0.8), Inches(0.55),
                font_size=13, bold=False, color=DEEP_BLUE)

    sep = add_shape_rect(slide, Inches(0.4), Inches(1.15), W - Inches(0.8), Pt(1), GRAY)

    txBox = slide.shapes.add_textbox(Inches(0.4), Inches(1.3), W - Inches(0.7),
                                      H - Inches(1.6))
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, bullet in enumerate(d["body"]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = p.add_run()
        run.text = bullet
        run.font.size = Pt(11)
        run.font.bold = False
        run.font.color.rgb = BODY_BLACK


def build_end(prs, d):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    W, H = prs.slide_width, prs.slide_height
    add_shape_rect(slide, 0, 0, W, H, DEEP_BLUE)
    add_textbox(slide, "휴먼컨설팅그룹",
                0, H / 2 - Inches(0.4), W, Inches(0.6),
                font_size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


# ─────────────────────────────────────────────
# 메인
# ─────────────────────────────────────────────
def main():
    OUTPUT = os.path.join(os.path.dirname(__file__), "HCG_Automated_Draft.pptx")

    prs = Presentation()
    prs.slide_width  = Inches(10.83)   # ★ 한국 표준 (v2 수정)
    prs.slide_height = Inches(7.50)

    print("HCG PPT 자동 생성 v2.0 시작...")
    print(f"슬라이드 크기: {prs.slide_width.inches:.2f}\" × {prs.slide_height.inches:.2f}\"")
    print(f"총 슬라이드 수: {len(SLIDES)}장\n")

    builders = {
        "cover":  build_cover,
        "toc":    build_toc,
        "normal": build_normal,
        "ref":    build_ref,
        "end":    build_end,
    }

    for d in SLIDES:
        builders[d["type"]](prs, d)
        print(f"  [완료] 슬라이드 {d['no']:02d}: {d['section']}")

    prs.save(OUTPUT)
    print(f"\n저장 완료: {OUTPUT}")
    print("PowerPoint에서 열어 디자인 세부 조정 후 사용하세요.")
    print("\n주요 변경 (v2.0):")
    print("  ✓ 슬라이드 크기: 13.33\" → 10.83\" (한국 표준)")
    print("  ✓ 본문 텍스트: Bold Blue → Regular Black")
    print("  ✓ 구조: _final 25장 구조 반영")
    print("  ✓ [참고] 슬라이드: 헤드 메시지 제거, 타이틀만")


if __name__ == "__main__":
    main()
