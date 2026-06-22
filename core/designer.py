# -*- coding: utf-8 -*-
"""
HCG 제안서 자동 PPT 생성기 v3.0
Design: 실제 final PPTX를 템플릿으로 사용 — 테마/폰트/레이아웃 100% 상속
Content: 롯데알미늄 직무기반 HR제도 설계 및 도입 제안서

핵심 변경 (v3.0):
  - 실제 PPTX 템플릿 기반 → 정확한 색상/폰트 자동 상속
  - Title placeholder → 자동 dark red (#921F0B 계열) 12pt bold
  - Subtitle placeholder (idx=10) → 자동 15pt 맑은 고딕 Semilight
  - 본문 item = noFill rounded rect, 0.5pt border, 맑은 고딕 11pt
  - 좌표 실측값 반영: items y=1.944~, h=0.496", w=4.528", gap=0.589"

실행: python auto_ppt.py
"""

import os, sys, math
sys.stdout.reconfigure(encoding='utf-8')

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_CONNECTOR
from pptx.oxml.ns import qn
from lxml import etree

# ── 경로 ──────────────────────────────────────────────────
BASE = ("C:\\Users\\cgpar\\OneDrive - 휴먼컨설팅그룹\\"
        "09 Admin\\09 etc\\other\\Claude\\롯데알미늄 제안서")
REAL = None  # template-path removed; DeckEngine.template param required
OUT  = None  # output-path removed; DeckEngine.out param required

# ── 추출된 실측 좌표 상수 ─────────────────────────────────
# 본문 slid e (layout 2 '본문')
TITLE_Y    = 0.250   # 제목 y (inch)
TITLE_X    = 0.575
TITLE_W    = 9.729
TITLE_H    = 0.305
SUB_Y      = 0.548   # 서브헤드 y
SUB_X      = 0.575
SUB_W      = 9.728
SUB_H      = 0.361

COL_L_X    = 0.691   # 좌측 컬럼 x
COL_R_X    = 5.615   # 우측 컬럼 x
ITEM_W     = 4.528   # 각 아이템 폭
ITEM_H     = 0.496   # 각 아이템 높이
ITEM_Y0    = 1.944   # 첫 번째 아이템 y 시작
ITEM_DY    = 0.589   # 아이템 간격 (y 피치)
HDR_Y      = 1.606   # 컬럼 헤더(OLE 대체) y

LABEL_X    = 0.369   # 행 라벨 x (추진 내용 등)
LABEL_W    = 1.250

# ── 색상 ──────────────────────────────────────────────────
HCG_RED    = RGBColor(0x92, 0x1F, 0x0B)
DARK_GRAY  = RGBColor(0x40, 0x40, 0x40)
MED_GRAY   = RGBColor(0x91, 0x91, 0x91)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
BLACK      = RGBColor(0x00, 0x00, 0x00)
# v4.1 고급 패턴용 (design 스킬 실측 팔레트)
GRAY_DK    = RGBColor(0x1D, 0x1D, 0x1D)
GRAY_LT    = RGBColor(0xD9, 0xD9, 0xD9)
WINE       = RGBColor(0x79, 0x40, 0x39)
CORAL_DK   = RGBColor(0x40, 0x0A, 0x07)
BLUE       = RGBColor(0x35, 0x6C, 0xB5)

# ── v6.0 KIA _final 실측 브랜드 팔레트 ────────────────────
#   (기아_중장기 보상체계 개선 추진_제안서_250912_v1.1_final.pptx 역추출)
#   FILL 빈도: E5838A(10) C72128(5) FFFFFF(5) E6E6E6 5D8ECF 3265A8 002459 D0E7F8...
#   FONT 빈도: 000000(1087) FFFFFF(86) F16249(56) C00000(18) 919191(5)
KIA_RED      = RGBColor(0xC7, 0x21, 0x28)   # HCG/기아 시그니처 레드 (활성 강조)
KIA_RED_LT   = RGBColor(0xE5, 0x83, 0x8A)   # 연한 레드 (비활성 마커)
ACCENT_OR    = RGBColor(0xF1, 0x62, 0x49)   # 강조 폰트 코랄
DEEP_RED     = RGBColor(0xC0, 0x00, 0x00)   # 진한 강조 레드
NAVY         = RGBColor(0x00, 0x24, 0x59)   # 차트 딥네이비
STEEL        = RGBColor(0x32, 0x65, 0xA8)   # 차트 스틸블루
SKY          = RGBColor(0x5D, 0x8E, 0xCF)   # 컨테이너 스카이블루
PALE_BLUE    = RGBColor(0xD0, 0xE7, 0xF8)   # 옅은 블루 배경
CONTAINER_GR = RGBColor(0xE6, 0xE6, 0xE6)   # 컨테이너 회색 배경
LINE_GRAY    = RGBColor(0x4F, 0x4F, 0x4F)   # 연결선 회색

# 기아 _final 템플릿 (best-practice 베이스). 없으면 REAL(롯데)로 폴백.
KIA_FINAL = None  # template-path removed; DeckEngine.template param required

# cm → inch 헬퍼 (실측 좌표가 cm 단위 → inch 변환)
def CM(v):
    """cm 값을 inch float로 변환 (1 inch = 2.54 cm)."""
    return v / 2.54


def to_rgb(c):
    """'#RRGGBB' / 'RRGGBB' 문자열 또는 RGBColor → RGBColor. None은 그대로."""
    if c is None or isinstance(c, RGBColor):
        return c
    s = str(c).lstrip('#')
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


# ── v7.0 파라다이스 _final 실측 럭셔리 골드 팔레트 ────────────
#   (HCG_파라다이스_보상제도 컨설팅_제안서_rev_final.pptx 역추출)
#   FILL: A49166/AC9A71(골드·샴페인) F1E5CA(크림) D4E1F2/AAC4E6/ECF6FC(블루)
#         6D8191(슬레이트). FONT 강조: C00000(딥레드) 356CB5(블루).
PARADISE_GOLD   = RGBColor(0xA4, 0x91, 0x66)   # 시그니처 골드(활성 accent)
PARADISE_GOLD2  = RGBColor(0xAC, 0x9A, 0x71)   # 보조 골드
PARADISE_CREAM  = RGBColor(0xF1, 0xE5, 0xCA)   # 크림 배경(컨테이너)
PARADISE_SLATE  = RGBColor(0x6D, 0x81, 0x91)   # 슬레이트 그레이
PARADISE_BLUE   = RGBColor(0xD4, 0xE1, 0xF2)   # 옅은 블루 배경
PARADISE_BLUE2  = RGBColor(0xAA, 0xC4, 0xE6)   # 블루 강조
PARADISE_EMPH   = RGBColor(0xC0, 0x00, 0x00)   # 폰트 강조 딥레드

# ── 테마 시스템 (v7.0) — 프로젝트별 브랜드 팔레트 전환 ────────
#   신규 도식 함수(create_toc_slide / add_structured_content_blocks /
#   add_container_box / content_blocks / 헤더·인사이트 헬퍼)가 THEME를
#   참조 → 한 줄로 전체 덱 브랜드 정합. 레드 2종 혼용 방지.
THEMES = {
    "hcg": {       # 기아/HCG 기본 (레드)
        "accent": RGBColor(0xC7, 0x21, 0x28),     # C72128
        "accent_lt": RGBColor(0xE5, 0x83, 0x8A),  # E5838A
        "container": RGBColor(0xE6, 0xE6, 0xE6),  # 회색
        "emphasis": RGBColor(0xC0, 0x00, 0x00),
        "ink": RGBColor(0x00, 0x00, 0x00),
        "on_accent": RGBColor(0xFF, 0xFF, 0xFF),
    },
    "paradise": {  # 파라다이스 (럭셔리 골드)
        "accent": PARADISE_GOLD,
        "accent_lt": PARADISE_GOLD2,
        "container": PARADISE_CREAM,
        "emphasis": PARADISE_EMPH,
        "ink": RGBColor(0x00, 0x00, 0x00),
        "on_accent": RGBColor(0xFF, 0xFF, 0xFF),
    },
}
# 어두운 배경(제목 흰색 처리)으로 간주할 색 집합 — 테마 무관 공통
DARK_FILLS = set()
THEME = dict(THEMES["hcg"])   # 현재 활성 테마 (기본 hcg)


def apply_theme(name):
    """활성 테마 전환. name in THEMES (모르면 hcg). 반환=테마 dict.
    전역 HCG_RED도 테마 accent로 재바인딩 → 기존 빌더(build_body_2col/
    process_roadmap/diff_matrix/add_table 등)가 자동으로 테마색 사용.
    (레거시 build()는 apply_theme 미호출 → 원래 921F0B 유지)."""
    global THEME, HCG_RED
    THEME = dict(THEMES.get(name, THEMES["hcg"]))
    HCG_RED = THEME["accent"]
    return THEME


def _dark_bg(color):
    """채움색이 어두워(텍스트 흰색 필요) 보이는지 판정 (상대휘도)."""
    c = to_rgb(color)
    if c is None:
        return False
    r, g, b = c[0], c[1], c[2]
    lum = 0.299 * r + 0.587 * g + 0.114 * b
    return lum < 140   # 140 미만이면 어두움 → 흰 텍스트


# ── 텍스트 오버플로 제어 (v7.0) ───────────────────────────────
def set_autofit_shrink(text_frame, font_scale=None, lnSpcReduction=None):
    """텍스트 프레임에 normAutofit(자동 축소) 적용 → 박스 넘침 방지.
    인간 _final이 noAutofit/spAutoFit를 544/452건 사용한 갭 보강.
    font_scale=None이면 PowerPoint가 열 때 자동 계산(권장)."""
    try:
        body = text_frame._txBody
    except Exception:
        return
    bodyPr = body.find(qn('a:bodyPr'))
    if bodyPr is None:
        bodyPr = body.makeelement(qn('a:bodyPr'), {})
        body.insert(0, bodyPr)
    # 기존 autofit 제거
    for tag in ('a:normAutofit', 'a:spAutoFit', 'a:noAutofit'):
        e = bodyPr.find(qn(tag))
        if e is not None:
            bodyPr.remove(e)
    na = etree.SubElement(bodyPr, qn('a:normAutofit'))
    if font_scale is not None:
        na.set('fontScale', str(int(font_scale * 1000)))
    if lnSpcReduction is not None:
        na.set('lnSpcReduction', str(int(lnSpcReduction * 1000)))
    return na


def fit_font_size(paragraphs_text, box_w_in, box_h_in, start_pt,
                  min_pt=7.0, ls=1.18, pad_in=0.10):
    """파이썬 측 동적 폰트 리사이징 — 추정 줄 수가 박스 높이를 넘으면
    start_pt부터 0.5pt씩 줄여 min_pt까지 맞춤. (한글 1자 ≈ font_pt 폭)
    paragraphs_text = [문단문자열,...]. 반환 = 결정된 pt."""
    usable_w_pt = max(1.0, (box_w_in - pad_in * 2) * 72.0)
    usable_h_pt = max(1.0, (box_h_in - pad_in * 2) * 72.0)
    fs = start_pt
    while fs > min_pt:
        cpl = max(1, int(usable_w_pt / (fs * 1.05)))   # 줄당 글자수
        lines = 0
        for t in paragraphs_text:
            L = len(t)
            lines += max(1, math.ceil(L / cpl)) if L else 1
        if lines * fs * ls <= usable_h_pt:
            break
        fs -= 0.5
    return round(max(fs, min_pt), 1)


# ════════════════════════════════════════════════════════════
# 헬퍼 함수
# ════════════════════════════════════════════════════════════

def _set_font_xml(run, name="맑은 고딕"):
    """run XML에 Korean 폰트 직접 설정"""
    rPr = run._r.get_or_add_rPr()
    for tag in [qn('a:latin'), qn('a:ea')]:
        el = rPr.find(tag)
        if el is None:
            el = etree.SubElement(rPr, tag)
        el.set('typeface', name)

def _set_no_fill(shape):
    """shape fill → noFill (transparent)"""
    sp_el = shape._element
    spPr = sp_el.find(qn('p:spPr'))
    if spPr is None:
        return
    ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    for tag in ['solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'noFill']:
        for el in spPr.findall(f'{{{ns}}}{tag}'):
            spPr.remove(el)
    etree.SubElement(spPr, f'{{{ns}}}noFill')

def _set_line_dark(shape, pt=0.5):
    """선 색상 = dk1(black), 두께 pt"""
    sp_el = shape._element
    spPr = sp_el.find(qn('p:spPr'))
    if spPr is None:
        return
    ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    ln_el = spPr.find(f'{{{ns}}}ln')
    if ln_el is None:
        ln_el = etree.SubElement(spPr, f'{{{ns}}}ln')
    ln_el.set('w', str(int(pt * 12700)))
    # color = dk1
    sF = etree.SubElement(ln_el, f'{{{ns}}}solidFill')
    sc = etree.SubElement(sF, f'{{{ns}}}schemeClr')
    sc.set('val', 'dk1')

def _add_run(para, text, fsize=11, bold=False, italic=False,
             color=None, font_name="맑은 고딕", spc=None, outline=None):
    run = para.add_run()
    run.text = text
    run.font.size = Pt(fsize)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    _set_font_xml(run, font_name)
    # 고급 타이포 (옵션, 기본 off → 기존 호출자 불변)
    if spc is not None:
        set_char_spacing(run, spc)
    if outline:
        if isinstance(outline, dict):
            set_text_outline(run, **outline)
        else:
            set_text_outline(run)
    return run

def _set_line_spacing(para, pct=112):
    """단락 행간 설정 (pct = 100~120...)"""
    pPr = para._p.get_or_add_pPr()
    ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    lnSpc = pPr.find(f'{{{ns}}}lnSpc')
    if lnSpc is not None:
        pPr.remove(lnSpc)
    lnSpc = etree.SubElement(pPr, f'{{{ns}}}lnSpc')
    spcPct = etree.SubElement(lnSpc, f'{{{ns}}}spcPct')
    spcPct.set('val', str(int(pct * 1000)))


# ── 슬라이드 삭제 ─────────────────────────────────────────
def delete_all_slides(prs):
    sldIdLst = prs.slides._sldIdLst
    for sld_el in list(sldIdLst):
        rId = sld_el.get(qn('r:id'))
        try:
            prs.part.drop_rel(rId)
        except Exception:
            pass
        sldIdLst.remove(sld_el)


# ── 본문 아이템 shape (rounded rect) ──────────────────────
def add_item(slide, text, left, top, width=ITEM_W, height=ITEM_H,
             fsize=11, bold=False, italic=False, align=PP_ALIGN.LEFT,
             fill_rgb=None, no_border=False, color=None,
             shadow=False, grad=None, anchor=None, spc=None, outline=None):
    """
    본문 콘텐츠 아이템 shape 추가
    - rounded rectangle (모서리 둥근 직사각형)
    - noFill (기본), 0.5pt dk1 border
    고급 옵션 (기본 off → 기존 호출자 불변):
      shadow  : True → outerShdw 그림자
      grad    : (color1, color2[, angle_deg]) → 그라디언트 채움 (fill_rgb 무시)
      anchor  : 't'/'ctr'/'b' 세로 정렬
      spc     : 자간 (OOXML spc 단위, 예 -30)
      outline : dict{color,width_pt} → 텍스트 윤곽선
    """
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )

    if grad:
        set_gradient(shape, *grad)
    elif fill_rgb:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_rgb
    else:
        _set_no_fill(shape)

    if no_border:
        # 선만 제거 (fill 유지 — _set_no_fill 호출 시 fill까지 사라지던 버그 수정)
        # Remove any line
        sp_el = shape._element
        spPr = sp_el.find(qn('p:spPr'))
        if spPr is not None:
            ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
            for ln_el in spPr.findall(f'{{{ns}}}ln'):
                spPr.remove(ln_el)
            noLn = etree.SubElement(spPr, f'{{{ns}}}ln')
            nf = etree.SubElement(noLn, f'{{{ns}}}noFill')
    else:
        _set_line_dark(shape, 0.5)

    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left  = Inches(0.07)
    tf.margin_right = Inches(0.04)
    tf.margin_top   = Inches(0.02)
    tf.margin_bottom = Inches(0.02)

    p = tf.paragraphs[0]
    p.alignment = align
    _add_run(p, text, fsize, bold, italic, color, spc=spc, outline=outline)
    _set_line_spacing(p, 112)

    if anchor:
        set_text_anchor(shape, anchor)
    if shadow:
        add_shadow(shape)

    return shape


# ── 텍스트 박스 ───────────────────────────────────────────
def add_textbox(slide, text, left, top, width, height,
                fsize=11, bold=False, italic=False, align=PP_ALIGN.LEFT,
                color=None, wrap=True):
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    _set_no_fill(txBox)
    sp_el = txBox._element
    spPr = sp_el.find(qn('p:spPr'))
    if spPr is not None:
        ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        ln_el = etree.SubElement(spPr, f'{{{ns}}}ln')
        nf = etree.SubElement(ln_el, f'{{{ns}}}noFill')

    tf = txBox.text_frame
    tf.word_wrap = wrap
    tf.margin_left = tf.margin_right = Inches(0.03)
    tf.margin_top = tf.margin_bottom = Inches(0.02)

    p = tf.paragraphs[0]
    p.alignment = align
    _add_run(p, text, fsize, bold, italic, color)
    _set_line_spacing(p, 112)
    return txBox


# ── 제목/서브헤드 설정 (placeholder 사용) ─────────────────
def set_title(slide, title_text, subtitle_text=None):
    """
    본문 레이아웃 placeholder[0] → 제목 (자동 dark red 스타일)
    placeholder[10] → 서브헤드 (자동 semilight 스타일)
    """
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:
            ph.text = title_text
        elif idx == 10 and subtitle_text:
            ph.text = subtitle_text


# ════════════════════════════════════════════════════════════
# 슬라이드 빌더
# ════════════════════════════════════════════════════════════

def build_cover(prs):
    """Slide 1: 표지 (layout 0)"""
    slide = prs.slides.add_slide(prs.slide_layouts[0])

    # Title placeholder (ctrTitle, idx=0)
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:
            ph.text = "롯데알미늄 직무기반HR 제도 설계 및 도입"
        elif idx == 1:
            ph.text = "- 제안서 -"

    # Date (별도 textbox — placeholder 아래)
    add_textbox(slide, "2026.03",
                1.137, 5.319, 8.563, 0.30,
                fsize=10, bold=False, align=PP_ALIGN.CENTER, color=DARK_GRAY)


def build_toc(prs):
    """Slide 2: 목차 (layout 1)"""
    slide = prs.slides.add_slide(prs.slide_layouts[1])

    # 목차 레이아웃 placeholder 확인
    ph_idxs = [ph.placeholder_format.idx for ph in slide.placeholders]

    # 목차 내용 — 번호 + 섹션명 + 슬라이드 범위
    toc_items = [
        ("I",   "Project Overview / 현황 진단",     "03-04"),
        ("II",  "직무체계 수립",                     "05-10"),
        ("III", "평가제도 개선",                     "11-15"),
        ("IV",  "보상제도 개선",                     "16-19"),
        ("V",   "추진계획 및 HCG 역량",             "20-23"),
    ]

    # 목차 타이틀
    add_textbox(slide, "목차 (Contents)",
                0.5, 0.3, 9.8, 0.5,
                fsize=18, bold=True, color=HCG_RED)

    # 구분선
    from pptx.util import Pt as _Pt
    line = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0.5), Inches(0.9), Inches(9.8), Inches(0.03)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = HCG_RED
    sp_el = line._element
    spPr = sp_el.find(qn('p:spPr'))
    if spPr is not None:
        ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        ln_el = etree.SubElement(spPr, f'{{{ns}}}ln')
        nf = etree.SubElement(ln_el, f'{{{ns}}}noFill')

    y = 1.1
    for i, (num, title, pg) in enumerate(toc_items):
        # 번호 원형
        circle = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.OVAL,
            Inches(0.5), Inches(y - 0.04), Inches(0.35), Inches(0.35)
        )
        circle.fill.solid()
        circle.fill.fore_color.rgb = HCG_RED
        circle.line.fill.background()
        ctf = circle.text_frame
        cp = ctf.paragraphs[0]
        cp.alignment = PP_ALIGN.CENTER
        _add_run(cp, num, fsize=9, bold=True, color=WHITE)

        # 섹션 타이틀
        add_textbox(slide, title,
                    1.0, y, 7.5, 0.4,
                    fsize=13, bold=False, color=DARK_GRAY)
        # 페이지 번호
        add_textbox(slide, pg,
                    8.8, y, 1.2, 0.4,
                    fsize=11, bold=False, align=PP_ALIGN.RIGHT, color=MED_GRAY)
        y += 0.85


def build_overview(prs):
    """Slide 3: Project Overview (layout 2)"""
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    set_title(slide,
              "Project Overview",
              "롯데알미늄의 직무체계 및 평가/보상제도 운영방안 도출")

    # 추진 내용 행
    add_textbox(slide, "추진 내용",
                LABEL_X, 2.175, LABEL_W, 0.32,
                fsize=13, italic=True, align=PP_ALIGN.RIGHT, color=MED_GRAY)
    add_textbox(slide, "롯데알미늄의 직무분류체계 수립, 직무기술서 작성(AI 활용), 평가/보상제도 설계를 통한 직무기반 HR제도 도입",
                1.836, 2.175, 8.759, 0.32,
                fsize=11, color=DARK_GRAY)

    # 추진 프로세스 행
    add_textbox(slide, "추진 프로세스",
                0.0, 2.952, 1.618, 0.32,
                fsize=13, italic=True, align=PP_ALIGN.RIGHT, color=MED_GRAY)

    steps = [
        ("현황 진단 및\n방향성 수립",  2.133),
        ("직무분류체계\n수립",          4.161),
        ("평가제도\n개선",              6.168),
        ("보상제도\n개선",              8.174),
    ]
    for title, x in steps:
        sh = add_item(slide, title, x, 2.952, width=1.976, height=0.666,
                      fsize=11, align=PP_ALIGN.CENTER,
                      fill_rgb=HCG_RED, color=WHITE)

    # 화살표 연결선 (간단한 텍스트 →)
    for x in [4.1, 6.1, 8.0]:
        add_textbox(slide, "→",
                    x, 3.18, 0.15, 0.25,
                    fsize=11, align=PP_ALIGN.CENTER, color=MED_GRAY)

    # 세부 설명
    sub_details = [
        (2.133, "현황 파악 Quick Review\n방향성 도출"),
        (4.161, "직무분류체계 초안\nAI 기반 직무기술서"),
        (6.168, "성과관리체계 수립\n평가 운영체계 개선"),
        (8.174, "보상 원칙/전략 수립\n보상 재원 Simulation"),
    ]
    for x, txt in sub_details:
        add_textbox(slide, txt, x, 3.632, 1.951, 0.554,
                    fsize=10, color=DARK_GRAY)

    # 추진 방안
    add_textbox(slide, "추진 방안",
                LABEL_X, 4.654, LABEL_W, 0.50,
                fsize=13, italic=True, align=PP_ALIGN.RIGHT, color=MED_GRAY)
    plan_text = ("추진 기간 : 12주\n"
                 "투입 인력 : PM 1명 + 컨설턴트 2명 (직무체계 전문 컨설턴트 투입)\n"
                 "추진 금액 : 225,000 천원 (부가세 별도; 세부 협의에 따라 변동 가능)")
    add_textbox(slide, plan_text,
                1.836, 4.589, 8.759, 0.972,
                fsize=10, color=DARK_GRAY)


def build_pain_point(prs):
    """Slide 4: Pain Point / 현황 진단 (layout 2) — 2-column"""
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    set_title(slide,
              "현황 진단 — Pain Point",
              "롯데알미늄 직무기반 HR제도 도입의 주요 Pain Point 및 개선 방향")

    # 컬럼 헤더
    hdr_items = [
        (COL_L_X, "현재 문제점 (As-Is)",   HCG_RED,   WHITE),
        (COL_R_X, "기대 개선사항 (To-Be)", DARK_GRAY, WHITE),
    ]
    for x, txt, fill, tcol in hdr_items:
        sh = add_item(slide, txt, x, HDR_Y, width=ITEM_W, height=0.333,
                      fsize=11, bold=True, align=PP_ALIGN.CENTER,
                      fill_rgb=fill, color=tcol, no_border=True)

    # Pain Point 데이터 (좌: 문제, 우: 개선)
    pairs = [
        ("직무 데이터 부재 — 체계적 직무 분류/기술 없이 관행적 운영",
         "직무분류체계 수립으로 명확한 직무 Scope 정의"),
        ("노후화된 직무 데이터 — 현업 실태 미반영, 활용 불가",
         "AI 활용 직무기술서 작성으로 신속한 현행화"),
        ("평가 기준 불명확 — 직무 특성 미반영, 공정성 낮음",
         "직무기반 성과관리체계 수립으로 평가 신뢰성 제고"),
        ("보상 운영 기준 부재 — 직무 가치 미반영, 형평성 문제",
         "직무가치 기반 보상체계 설계로 내외부 공정성 확보"),
        ("HR제도 활용 미흡 — 채용/배치/육성 연계 부족",
         "직무기반 통합 HR제도 운영체계 구축"),
    ]

    for i, (left_txt, right_txt) in enumerate(pairs):
        y = ITEM_Y0 + i * ITEM_DY
        add_item(slide, left_txt,  COL_L_X, y, fsize=10, color=DARK_GRAY)
        add_item(slide, right_txt, COL_R_X, y, fsize=10, color=DARK_GRAY)

    # 중간 VS 연결 표시
    for i in range(len(pairs)):
        y = ITEM_Y0 + i * ITEM_DY + 0.05
        add_item(slide, "→", 5.22, y, width=0.394, height=0.394,
                 fsize=10, bold=True, align=PP_ALIGN.CENTER,
                 fill_rgb=HCG_RED, color=WHITE, no_border=True)


def build_section_divider(prs, section_num, section_title, sub):
    """Section 구분 슬라이드 (layout 1 '목차')"""
    slide = prs.slides.add_slide(prs.slide_layouts[1])

    # 섹션 번호 원형
    circle = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.OVAL,
        Inches(1.0), Inches(2.8), Inches(0.9), Inches(0.9)
    )
    circle.fill.solid()
    circle.fill.fore_color.rgb = HCG_RED
    circle.line.fill.background()
    ctf = circle.text_frame
    cp = ctf.paragraphs[0]
    cp.alignment = PP_ALIGN.CENTER
    _add_run(cp, section_num, fsize=22, bold=True, color=WHITE)

    # 섹션 타이틀
    add_textbox(slide, section_title,
                2.2, 2.7, 8.0, 0.7,
                fsize=24, bold=True, color=HCG_RED)

    # 서브 설명
    add_textbox(slide, sub,
                2.2, 3.5, 8.0, 0.5,
                fsize=13, color=DARK_GRAY)

    # 구분선
    line = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(2.2), Inches(3.42), Inches(8.0), Inches(0.03)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = MED_GRAY
    sp_el = line._element
    spPr = sp_el.find(qn('p:spPr'))
    if spPr is not None:
        ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        ln_el = etree.SubElement(spPr, f'{{{ns}}}ln')
        nf = etree.SubElement(ln_el, f'{{{ns}}}noFill')


def build_body_2col(prs, title, subtitle, header_l, header_r, left_items, right_items):
    """표준 2컬럼 본문 슬라이드"""
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    set_title(slide, title, subtitle)

    # 컬럼 헤더
    add_item(slide, header_l, COL_L_X, HDR_Y, width=ITEM_W, height=0.333,
             fsize=11, bold=True, align=PP_ALIGN.CENTER,
             fill_rgb=HCG_RED, color=WHITE, no_border=True)
    add_item(slide, header_r, COL_R_X, HDR_Y, width=ITEM_W, height=0.333,
             fsize=11, bold=True, align=PP_ALIGN.CENTER,
             fill_rgb=DARK_GRAY, color=WHITE, no_border=True)

    rows = max(len(left_items), len(right_items))
    for i in range(rows):
        y = ITEM_Y0 + i * ITEM_DY
        if i < len(left_items):
            add_item(slide, left_items[i],  COL_L_X, y, fsize=10, color=DARK_GRAY)
        if i < len(right_items):
            add_item(slide, right_items[i], COL_R_X, y, fsize=10, color=DARK_GRAY)

    return slide


def build_body_single(prs, title, subtitle, rows):
    """단일 컬럼 본문 슬라이드 (label + content 구조)"""
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    set_title(slide, title, subtitle)

    y = ITEM_Y0
    for label, content in rows:
        if label:
            add_textbox(slide, label,
                        LABEL_X, y, LABEL_W, ITEM_H,
                        fsize=12, italic=True, align=PP_ALIGN.RIGHT, color=MED_GRAY)
        add_item(slide, content,
                 1.836, y, width=8.759, height=ITEM_H,
                 fsize=10, color=DARK_GRAY)
        y += ITEM_DY

    return slide


def build_body_process(prs, title, subtitle, steps, desc_rows=None):
    """프로세스 단계 슬라이드 (→ 화살표 구조)"""
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    set_title(slide, title, subtitle)

    step_y = 2.5
    n = len(steps)
    box_w = min(2.2, 9.0 / n)
    start_x = 0.5 + (9.5 - n * box_w - (n-1) * 0.25) / 2

    for i, (step_title, step_detail) in enumerate(steps):
        x = start_x + i * (box_w + 0.25)
        sh = add_item(slide, step_title, x, step_y, width=box_w, height=0.6,
                      fsize=11, bold=True, align=PP_ALIGN.CENTER,
                      fill_rgb=HCG_RED, color=WHITE, no_border=True)
        if step_detail:
            add_textbox(slide, step_detail, x, step_y + 0.65, box_w, 0.8,
                        fsize=9, align=PP_ALIGN.CENTER, color=DARK_GRAY)
        if i < n - 1:
            add_textbox(slide, "→",
                        x + box_w + 0.03, step_y + 0.15, 0.2, 0.3,
                        fsize=13, align=PP_ALIGN.CENTER, color=MED_GRAY)

    if desc_rows:
        y = step_y + 1.6
        for label, content in desc_rows:
            if label:
                add_textbox(slide, label,
                            LABEL_X, y, LABEL_W, 0.38,
                            fsize=11, italic=True, align=PP_ALIGN.RIGHT, color=MED_GRAY)
            add_item(slide, content,
                     1.836, y, width=8.759, height=0.4,
                     fsize=10, color=DARK_GRAY)
            y += 0.46

    return slide


def build_end(prs):
    """Slide 24: End of document (layout 3)"""
    slide = prs.slides.add_slide(prs.slide_layouts[3])
    # 레이아웃이 제공하는 디자인 사용
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:
            try: ph.text = "Thank you"
            except: pass
    return slide


def build_appendix(prs):
    """Slide 25: Appendix"""
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    set_title(slide, "Appendix", "직무분류체계 예시 — 대기업 알미늄 제조업 기준")

    rows = [
        ("대분류", "영업/마케팅  |  생산/제조  |  품질/안전  |  연구/개발  |  경영지원  |  구매/물류"),
        ("중분류",
         "영업관리, 마케팅전략, 고객지원 / 판재생산, 박재생산, 주조 / 품질관리, 안전환경 / 연구개발, 공정개발"),
        ("소분류",
         "영업기획, 국내영업, 해외영업, 마케팅조사, 고객불만처리 / 판재제조, 박재제조, 용해주조, 열처리 / ..."),
        ("직무기술서",
         "각 직무별 Key Responsibilities, Required Competency, Required Knowledge/Skill 정의 (AI 초안 기반)"),
    ]

    y = ITEM_Y0
    for label, content in rows:
        add_textbox(slide, label,
                    LABEL_X, y, LABEL_W, 0.45,
                    fsize=12, italic=True, align=PP_ALIGN.RIGHT, color=MED_GRAY)
        add_item(slide, content,
                 1.836, y, width=8.759, height=0.45,
                 fsize=10, color=DARK_GRAY)
        y += 0.55


# ════════════════════════════════════════════════════════════
# 메인 빌드
# ════════════════════════════════════════════════════════════

# ════════════════════════════════════════════════════════════
# 고급 도식 헬퍼 (v4.1 — 롯데알미늄 _final 실측 패턴)
# ════════════════════════════════════════════════════════════

def add_header_bar(slide, text, x=0.691, y=1.595, w=9.451, h=0.333, fill=None):
    """콘텐츠 영역 헤더 바 (GRAY fill + 흰 bold)."""
    return add_item(slide, text, x, y, width=w, height=h, fsize=11, bold=True,
                    align=PP_ALIGN.LEFT, fill_rgb=fill or MED_GRAY, color=WHITE, no_border=True)

def add_label_tag(slide, text, x, y, w=1.02, h=0.5, fill=None):
    """미니 카테고리 라벨 태그."""
    return add_item(slide, text, x, y, width=w, height=h, fsize=9, bold=True,
                    align=PP_ALIGN.CENTER, fill_rgb=fill or GRAY_DK, color=WHITE, no_border=True)

def add_connector_badge(slide, text, x=4.752, y=1.994, w=1.25, h=0.39, fill=None):
    """중앙 커넥터/키워드 배지."""
    return add_item(slide, text, x, y, width=w, height=h, fsize=9, bold=True,
                    align=PP_ALIGN.CENTER, fill_rgb=fill or MED_GRAY, color=WHITE, no_border=True)

def add_vs_badge(slide, x=5.22, y=1.61, w=0.39, h=0.39, fill=None, text="VS"):
    """Legacy vs HCG 'VS' 원형 배지."""
    sh = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL,
                                Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill or HCG_RED
    sh.line.fill.background()
    p = sh.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    _add_run(p, text, fsize=10, bold=True, color=WHITE)
    return sh

def add_example_badge(slide, x=9.56, y=1.61, w=0.58, h=0.27):
    """'예시적' 우상단 배지 — 가설/예시 데이터 단정 회피."""
    return add_item(slide, "예시적", x, y, width=w, height=h, fsize=8,
                    align=PP_ALIGN.CENTER, fill_rgb=GRAY_LT, color=DARK_GRAY, no_border=True)

def add_insight_quote(slide, text, x=0.691, y=6.79, w=9.451, h=0.30, color=None):
    """큰따옴표 italic 강조 takeaway."""
    if not (text.startswith("“") or text.startswith('"')):
        text = "“" + text + "”"
    return add_textbox(slide, text, x, y, w, h, fsize=11, bold=True, italic=True,
                       align=PP_ALIGN.CENTER, color=color or THEME["accent"])


def build_overview_3col(prs, title, subtitle, cols, bar_label=None, example=False):
    """모듈 Overview — Step1/2/3 3컬럼 + AI 콜아웃. cols=[{header,desc,ai,detail}]×3"""
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    set_title(slide, title, subtitle)
    if bar_label:
        add_header_bar(slide, bar_label)
    xs = [0.69, 3.90, 7.11]; w = 3.03
    for i, col in enumerate(cols[:3]):
        x = xs[i]
        add_item(slide, "", x, 2.88, width=w, height=4.17, fill_rgb=GRAY_LT, no_border=True)
        add_item(slide, col.get("header", ""), x, 2.10, width=w, height=0.79, fsize=11, bold=True,
                 align=PP_ALIGN.CENTER, fill_rgb=HCG_RED, color=WHITE, no_border=True)
        add_textbox(slide, col.get("desc", ""), x + 0.10, 2.98, w - 0.20, 0.80, fsize=9, color=DARK_GRAY)
        add_item(slide, col.get("ai", ""), x + 0.10, 3.86, width=w - 0.20, height=1.45, fsize=9, bold=True,
                 align=PP_ALIGN.CENTER, fill_rgb=BLUE, color=WHITE, no_border=True)
        add_textbox(slide, col.get("detail", ""), x + 0.10, 5.45, w - 0.20, 1.45, fsize=8.5, color=DARK_GRAY)
    if example:
        add_example_badge(slide)
    return slide


def build_diff_matrix(prs, title, subtitle, rows, bottom_quote=None, example=False):
    """직군별 차별화 매트릭스. rows=[{group,trait,insight,apply}] (최대 4)"""
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    set_title(slide, title, subtitle)
    add_textbox(slide, "업무 특성", 2.00, 2.05, 2.07, 0.30, fsize=9, bold=True, align=PP_ALIGN.CENTER, color=MED_GRAY)
    add_textbox(slide, "핵심 방향성", 4.21, 2.05, 2.68, 0.30, fsize=9, bold=True, align=PP_ALIGN.CENTER, color=MED_GRAY)
    add_textbox(slide, "평가 반영방안", 7.57, 2.05, 2.59, 0.30, fsize=9, bold=True, align=PP_ALIGN.CENTER, color=MED_GRAY)
    ys = [2.46, 3.55, 4.63, 5.74]
    for i, r in enumerate(rows[:4]):
        y = ys[i]
        add_item(slide, r.get("group", ""), 0.69, y, width=1.25, height=0.97, fsize=10, bold=True,
                 align=PP_ALIGN.CENTER, fill_rgb=HCG_RED, color=WHITE, no_border=True)
        add_textbox(slide, r.get("trait", ""), 2.00, y, 2.07, 0.97, fsize=8.5, color=DARK_GRAY)
        ins = r.get("insight", "")
        if not ins.startswith("“"):
            ins = "“" + ins + "”"
        add_item(slide, ins, 4.21, y, width=2.68, height=0.97, fsize=9, bold=True,
                 align=PP_ALIGN.CENTER, fill_rgb=GRAY_LT, color=HCG_RED)
        add_textbox(slide, "▶", 6.99, y + 0.30, 0.47, 0.35, fsize=12, align=PP_ALIGN.CENTER, color=MED_GRAY)
        add_textbox(slide, r.get("apply", ""), 7.57, y, 2.59, 0.97, fsize=8.5, color=DARK_GRAY)
    if bottom_quote:
        add_insight_quote(slide, bottom_quote, 0.69, 6.82, 9.45)
    if example:
        add_example_badge(slide)
    return slide


def build_pain_point_categorized(prs, title, subtitle, left_hdr, right_hdr, rows,
                                 summary_left=None, summary_right=None):
    """Pain Point 카테고리화. rows=[(left_label,left_text,mid,right_text)] (최대 6)"""
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    set_title(slide, title, subtitle)
    add_header_bar(slide, left_hdr, 0.691, 1.61, 4.528, 0.333)
    add_header_bar(slide, right_hdr, 5.615, 1.61, 4.528, 0.333)
    y0, dy, h = 1.99, 0.585, 0.50
    for i, row in enumerate(rows[:6]):
        llab, ltext, mid, rtext = row
        y = y0 + i * dy
        add_item(slide, ltext, 1.78, y, width=3.44, height=h, fsize=9, color=DARK_GRAY)
        add_label_tag(slide, llab, 0.691, y, 1.02, h, fill=GRAY_DK)
        if mid:
            add_connector_badge(slide, mid, 4.752, y + 0.05, 1.25, 0.39, fill=MED_GRAY)
        add_item(slide, rtext, 5.615, y, width=4.528, height=h, fsize=9, color=DARK_GRAY)
    if summary_left:
        add_item(slide, summary_left, 0.691, 5.62, width=4.528, height=0.92, fsize=10, bold=True,
                 align=PP_ALIGN.CENTER, fill_rgb=GRAY_LT, color=DARK_GRAY, no_border=True)
    if summary_right:
        add_item(slide, summary_right, 5.615, 5.62, width=4.528, height=0.92, fsize=10, bold=True,
                 align=PP_ALIGN.CENTER, fill_rgb=HCG_RED, color=WHITE, no_border=True)
    return slide


def build_approach_vs(prs, title, subtitle, left_title, left_items, right_title, right_items, bottom_quote=None):
    """Legacy vs HCG 'VS' 대비 블록 + 하단 takeaway."""
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    set_title(slide, title, subtitle)
    add_item(slide, "", 0.69, 2.20, width=4.52, height=4.10, fill_rgb=GRAY_LT, no_border=True)
    add_textbox(slide, left_title, 0.80, 2.30, 4.30, 0.40, fsize=12, bold=True, color=DARK_GRAY)
    yy = 2.92
    for it in left_items:
        add_textbox(slide, "• " + it, 0.92, yy, 4.08, 0.70, fsize=10, color=DARK_GRAY); yy += 0.74
    add_item(slide, "", 5.65, 2.20, width=4.45, height=4.10, fill_rgb=None, no_border=False)
    add_textbox(slide, right_title, 5.78, 2.30, 4.20, 0.40, fsize=12, bold=True, color=HCG_RED)
    yy = 2.92
    for it in right_items:
        add_textbox(slide, "• " + it, 5.88, yy, 4.05, 0.70, fsize=10, color=DARK_GRAY); yy += 0.74
    add_vs_badge(slide, 5.22, 3.95, 0.39, 0.39, fill=HCG_RED, text="VS")
    if bottom_quote:
        add_insight_quote(slide, bottom_quote, 1.10, 6.60, 8.64)
    return slide


# ════════════════════════════════════════════════════════════
# 고급 도식 헬퍼 v2 (v4.2 — 기아 _final 비교: 도식화/밀도 강화)
# ════════════════════════════════════════════════════════════

def add_connector_line(slide, x1, y1, x2, y2, color=None, w_pt=1.0):
    """박스 간 직선 연결선 (흐름·관계 시각화)."""
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                      Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    conn.line.color.rgb = color or MED_GRAY
    conn.line.width = Pt(w_pt)
    return conn

def add_arrow_flow(slide, centers, color=None):
    """중심점 리스트 사이에 ▶ 화살표 자동 배치."""
    for (x1, y1), (x2, y2) in zip(centers, centers[1:]):
        mx, my = (x1 + x2) / 2.0, (y1 + y2) / 2.0
        add_textbox(slide, "▶", mx - 0.12, my - 0.15, 0.26, 0.30,
                    fsize=12, align=PP_ALIGN.CENTER, color=color or MED_GRAY)

def add_icon_placeholder(slide, x, y, d=0.5, fill=None):
    """아이콘 대체 원형 도형 (PICTURE 자동생성 불가 대응)."""
    sh = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL,
                                Inches(x), Inches(y), Inches(d), Inches(d))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill or BLUE
    sh.line.fill.background()
    return sh

def add_table(slide, x, y, w, h, data, header=True, fsize=9):
    """비교용 TABLE. data=2차원 리스트. header행 HCG_RED·흰 bold."""
    rows, cols = len(data), len(data[0])
    gf = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(h))
    tbl = gf.table
    for r in range(rows):
        for c in range(cols):
            cell = tbl.cell(r, c)
            cell.text = str(data[r][c])
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                for run in p.runs:
                    run.font.size = Pt(fsize)
                    _set_font_xml(run)
                    if header and r == 0:
                        run.font.bold = True
                        run.font.color.rgb = WHITE
            cell.fill.solid()
            cell.fill.fore_color.rgb = HCG_RED if (header and r == 0) else WHITE
    return gf

def build_process_roadmap(prs, title, subtitle, phases):
    """Process Overview 로드맵 — phase 박스 + 하위 step + ▶ 연결. phases=[{name,steps:[...]}]"""
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    set_title(slide, title, subtitle)
    n = len(phases); gap = 0.30; total_w = 9.451; x0 = 0.691
    pw = (total_w - gap * (n - 1)) / n
    add_connector_line(slide, x0, 2.41, x0 + total_w, 2.41, color=MED_GRAY, w_pt=1.5)  # phase baseline
    for i, ph in enumerate(phases):
        x = x0 + i * (pw + gap)
        add_item(slide, ph.get("name", ""), x, 2.10, width=pw, height=0.62, fsize=11, bold=True,
                 align=PP_ALIGN.CENTER, fill_rgb=HCG_RED, color=WHITE, no_border=True)
        sy = 3.00
        for st in ph.get("steps", []):
            add_item(slide, st, x, sy, width=pw, height=0.66, fsize=9, align=PP_ALIGN.CENTER,
                     fill_rgb=GRAY_LT, color=DARK_GRAY, no_border=True)
            sy += 0.74
        if i < n - 1:
            ax = x + pw + (gap - 0.26) / 2.0
            add_textbox(slide, "▶", ax, 2.22, 0.26, 0.34, fsize=14, align=PP_ALIGN.CENTER, color=MED_GRAY)
    return slide

def build_compare_table(prs, title, subtitle, headers, rows, example=False):
    """표 기반 비교 장표. headers=리스트, rows=[[...]]."""
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    set_title(slide, title, subtitle)
    data = [headers] + rows
    h = min(4.6, 0.55 * len(data) + 0.2)
    add_table(slide, 0.691, 2.05, 9.451, h, data, header=True)
    if example:
        add_example_badge(slide)
    return slide


# ════════════════════════════════════════════════════════════
# 고급 XML 레벨 디자인 헬퍼 (v5.0 — 인간 _final XML 딥다이브 기반)
#   인간 _final 실측: 텍스트윤곽선 1807건, 그림자 60, 투명도 2478,
#   그라디언트 43, 자간 -30~-100. AI Draft 전부 0 → 갭 보강.
# ════════════════════════════════════════════════════════════

A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

def _A(tag):
    return f'{{{A_NS}}}{tag}'

def _hex(color):
    """RGBColor 또는 hex 문자열 → 6자리 hex 문자열."""
    s = str(color)
    return s[-6:].upper() if len(s) >= 6 else s.upper()

def _spPr(shape):
    """shape의 p:spPr element 반환 (없으면 None)."""
    return shape._element.find(qn('p:spPr'))

def set_char_spacing(run, val):
    """자간(kerning) 설정. val = OOXML spc 단위(1/100 pt). 음수=좁힘(예 -30).
    인간 _final 본문 자간 -30, 제목 -50~-100 패턴."""
    rPr = run._r.get_or_add_rPr()
    rPr.set('spc', str(int(val)))
    return rPr

def set_text_outline(run, color="FFFFFF", width_pt=0.75):
    """텍스트 윤곽선 <a:rPr><a:ln>. 인간 _final 최다 패턴(1807건) —
    채움색과 윤곽선 대비로 가독성·강조 확보. ln은 rPr 첫 자식이어야 함."""
    rPr = run._r.get_or_add_rPr()
    old = rPr.find(_A('ln'))
    if old is not None:
        rPr.remove(old)
    ln = rPr.makeelement(_A('ln'), {})
    ln.set('w', str(int(width_pt * 12700)))
    sf = etree.SubElement(ln, _A('solidFill'))
    clr = etree.SubElement(sf, _A('srgbClr'))
    clr.set('val', _hex(color))
    rPr.insert(0, ln)
    return ln

def add_shadow(shape, blur_pt=4.0, dist_pt=2.5, direction=2700000,
               alpha_pct=55, color="000000"):
    """outerShdw 그림자. spPr 자식 순서상 effectLst는 ln 뒤 → append.
    alpha_pct = 그림자 자체 투명도(%) (55 → 약간 옅은 그림자)."""
    spPr = _spPr(shape)
    if spPr is None:
        return None
    old = spPr.find(_A('effectLst'))
    if old is not None:
        spPr.remove(old)
    eff = etree.SubElement(spPr, _A('effectLst'))
    sh = etree.SubElement(eff, _A('outerShdw'))
    sh.set('blurRad', str(int(blur_pt * 12700)))
    sh.set('dist', str(int(dist_pt * 12700)))
    sh.set('dir', str(int(direction)))
    sh.set('rotWithShape', '0')
    clr = etree.SubElement(sh, _A('srgbClr'))
    clr.set('val', _hex(color))
    al = etree.SubElement(clr, _A('alpha'))
    al.set('val', str(int(alpha_pct * 1000)))
    return eff

def set_transparency(shape, pct):
    """solidFill 채움 투명도 pct%(0=불투명, 100=완전투명)."""
    spPr = _spPr(shape)
    if spPr is None:
        return
    sf = spPr.find(_A('solidFill'))
    if sf is None:
        return
    clr = sf.find(_A('srgbClr'))
    if clr is None:
        clr = sf.find(_A('schemeClr'))
    if clr is None:
        return
    for a in clr.findall(_A('alpha')):
        clr.remove(a)
    al = etree.SubElement(clr, _A('alpha'))
    al.set('val', str(int((100 - pct) * 1000)))

def set_gradient(shape, color1, color2, angle_deg=90):
    """선형 그라디언트 채움. color1(0%)→color2(100%), angle_deg(시계방향).
    spPr fill 위치(prstGeom 뒤, ln 앞)에 삽입."""
    spPr = _spPr(shape)
    if spPr is None:
        return
    for tag in ['solidFill', 'gradFill', 'noFill', 'blipFill', 'pattFill', 'grpFill']:
        for el in spPr.findall(_A(tag)):
            spPr.remove(el)
    grad = etree.Element(_A('gradFill'))
    gs = etree.SubElement(grad, _A('gsLst'))
    for pos, c in [(0, color1), (100000, color2)]:
        g = etree.SubElement(gs, _A('gs'))
        g.set('pos', str(pos))
        cc = etree.SubElement(g, _A('srgbClr'))
        cc.set('val', _hex(c))
    lin = etree.SubElement(grad, _A('lin'))
    lin.set('ang', str(int(angle_deg * 60000)))
    lin.set('scaled', '1')
    ln = spPr.find(_A('ln'))
    if ln is not None:
        ln.addprevious(grad)
    else:
        spPr.append(grad)

def set_text_anchor(shape, anchor='ctr'):
    """텍스트 세로 정렬: 't'(상), 'ctr'(중앙), 'b'(하)."""
    try:
        bodyPr = shape.text_frame._txBody.find(qn('a:bodyPr'))
        if bodyPr is not None:
            bodyPr.set('anchor', anchor)
    except Exception:
        pass


# ════════════════════════════════════════════════════════════
# Phase 2/3 핵심 — Text-to-Visual 구조화 엔진 (v6.0)
#   인간 _final 역추출 패턴을 그대로 재현하는 좌표·도형 함수.
#   ① create_toc_slide        : 목차=텍스트나열 폐기 → 인간 어젠다카드 재현
#   ② add_structured_content_blocks : 리스트 텍스트 → N등분 카드 그리드
#   ③ add_container_box        : 박스 안 박스(컨테이너) 레이아웃
# ════════════════════════════════════════════════════════════

# ── 목차/섹션 어젠다 카드 (인간 _final slide 1·17·52 실측 재현) ──
#   구조(cm): 외곽카드 x5.4 y4.93 w20.37 h(=N행). 행 h=3.54 contiguous.
#   행 = [좌측 accent블록 w2.74 | 로마숫자 | 섹션명 18pt]. 활성 행만
#   accent=KIA_RED(C72128)+빨간 윤곽 하이라이트, 비활성 accent=KIA_RED_LT.
ROMAN = ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ"]


def create_toc_slide(prs, sections, current=None, layout_idx=1,
                     title=None):
    """목차/섹션 디바이더 슬라이드 생성 (인간 _final 어젠다카드 100% 재현).

    sections : ["제안 배경", "수행 방안", "일정 및 조직", ...]  (섹션명 리스트)
    current  : 강조할 섹션 인덱스(0-base). None이면 전체 목차(첫 섹션 강조).
    layout_idx : 사용할 레이아웃 (1='목차').
    title    : 좌상단 제목(옵션). None이면 생략(인간 final은 제목 없음).

    좌표는 실측 cm값을 CM()으로 inch 변환 — 인간본과 픽셀 단위 정합.
    """
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    n = len(sections)

    # 외곽 카드 (white solid, 0.5pt 테두리)
    card_x, card_w = CM(5.4), CM(20.37)
    row_h = CM(3.54)
    total_h = row_h * n
    card_y = (DeckEngine.SLIDE_H - total_h) / 2.0 + 0.30   # 살짝 아래 정렬
    card = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE,
                                  Inches(card_x), Inches(card_y - CM(0.64)),
                                  Inches(card_w), Inches(total_h + CM(1.28)))
    card.fill.solid(); card.fill.fore_color.rgb = WHITE
    _set_line_dark(card, 0.5)
    add_shadow(card, blur_pt=6, dist_pt=3, alpha_pct=70)

    accent_x, accent_w = CM(5.79), CM(2.74)
    roman_x, roman_w = CM(6.77), CM(0.79)
    title_x = CM(10.54)
    title_w = CM(12.93)

    if current is None:
        current = 0

    for i, sec in enumerate(sections):
        y = card_y + i * row_h

        # 행 배경 (white, 옅은 테두리) — contiguous 행 구분선 효과
        row_bg = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE,
                                        Inches(accent_x), Inches(y),
                                        Inches(card_w - CM(0.78)), Inches(row_h))
        row_bg.fill.solid(); row_bg.fill.fore_color.rgb = WHITE
        _set_line_dark(row_bg, 0.5)

        # 좌측 accent 블록 (활성=THEME accent, 비활성=accent_lt)
        active = (i == current)
        accent = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE,
                                        Inches(accent_x), Inches(y),
                                        Inches(accent_w), Inches(row_h))
        accent.fill.solid()
        accent.fill.fore_color.rgb = THEME["accent"] if active else THEME["accent_lt"]
        accent.line.fill.background()

        # 로마 숫자 (accent 위, 18pt bold 흰색)
        rbox = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE,
                                      Inches(accent_x), Inches(y),
                                      Inches(accent_w), Inches(row_h))
        rbox.fill.background(); rbox.line.fill.background()
        set_text_anchor(rbox, 'ctr')
        rp = rbox.text_frame.paragraphs[0]; rp.alignment = PP_ALIGN.CENTER
        _add_run(rp, ROMAN[i] if i < len(ROMAN) else str(i + 1),
                 fsize=18, bold=True, color=WHITE)

        # 섹션명 (18pt, 활성=THEME accent / 비활성=BLACK)
        tbox = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE,
                                      Inches(title_x), Inches(y),
                                      Inches(title_w), Inches(row_h))
        tbox.fill.background(); tbox.line.fill.background()
        set_text_anchor(tbox, 'ctr')
        tp = tbox.text_frame.paragraphs[0]; tp.alignment = PP_ALIGN.LEFT
        _add_run(tp, sec, fsize=18, bold=active, spc=-30,
                 color=THEME["accent"] if active else BLACK)

        # 활성 행 윤곽 하이라이트 (인간 final FREEFORM 재현, 테마 accent)
        if active:
            hl = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE,
                                        Inches(accent_x + CM(0.37)),
                                        Inches(y + CM(0.33)),
                                        Inches(CM(19.61)), Inches(CM(3.31) * row_h / CM(3.54)))
            hl.fill.background()
            hl.line.color.rgb = THEME["accent"]
            hl.line.width = Pt(1.5)

    if title:
        add_textbox(slide, title, CM(1.46), CM(0.64), CM(24.71), CM(0.78),
                    fsize=18, bold=True, color=THEME["accent"])
    return slide


# ── 구조화 콘텐츠 블록: 리스트 텍스트 → N등분 카드 그리드 ──
#   Phase 2 계산식: 가용폭 W, n개 카드, 카드간 gap → card_w=(W-gap*(n-1))/n.
#   각 카드 = 모서리 둥근 직사각형 컨테이너 + 상단 헤더밴드(accent) + 본문.
#   margin/padding 내장, 윤곽선·그림자 XML 적용 (인간본 디자인 밀도).
def add_structured_content_blocks(slide, items, x=None, y=2.05,
                                   w=None, h=4.30, gap=0.26,
                                   accent=None, header_h=0.62,
                                   pad=0.13, title_size=12,
                                   body_size=10.5, numbered=True,
                                   body_align=PP_ALIGN.LEFT,
                                   shadow=True):
    """리스트형 텍스트를 자동으로 N개의 카드 그리드로 변환 배치.

    items : [{"title": "...", "body": "...", "bullets": [..](옵션)}, ...]
            (문자열만 주면 title 없는 단일 본문 카드로 처리)
    x, w  : 그리드 영역 좌측 x / 전체 폭 (None이면 본문 표준 마진 사용)
    n     : len(items) (최대 4 권장; 그 이상은 자동 축소 폰트)
    반환  : 생성된 카드 컨테이너 shape 리스트.

    계산식:
      card_w = (w - gap*(n-1)) / n
      카드 내부 본문폭 = card_w - 2*pad
      헤더밴드 = 카드 상단 header_h, accent 채움 + 흰 bold 타이틀
      본문영역 = 헤더 아래 (h - header_h - pad)
    """
    if x is None:
        x = CM(1.46)               # 본문 좌측 마진 실측
    if w is None:
        w = CM(24.71)              # 본문 가용폭 실측
    accent = to_rgb(accent) if accent is not None else THEME["accent"]
    n = max(1, len(items))
    card_w = (w - gap * (n - 1)) / n
    # 카드 수 많으면 폰트 자동 축소
    if n >= 4:
        title_size = min(title_size, 11)
        body_size = min(body_size, 9.5)
    if n >= 5:
        title_size = min(title_size, 10)
        body_size = min(body_size, 9)

    shapes = []
    for i, item in enumerate(items):
        if isinstance(item, str):
            item = {"title": "", "body": item}
        cx = x + i * (card_w + gap)

        # ① 카드 컨테이너 (모서리 둥근 직사각형, 옅은 회색 배경 + 그림자)
        card = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
                                      Inches(cx), Inches(y),
                                      Inches(card_w), Inches(h))
        card.fill.solid(); card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = accent
        card.line.width = Pt(1.0)
        if shadow:
            add_shadow(card, blur_pt=5, dist_pt=2.5, alpha_pct=68)
        shapes.append(card)

        # ② 헤더 밴드 (accent 채움, 흰 bold 타이틀) — 자간·윤곽선 적용
        title = item.get("title", "")
        if title:
            head_txt = (f"{i+1}. {title}" if numbered else title)
            hb = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
                                        Inches(cx), Inches(y),
                                        Inches(card_w), Inches(header_h))
            hb.fill.solid(); hb.fill.fore_color.rgb = accent
            hb.line.fill.background()
            set_text_anchor(hb, 'ctr')
            hp = hb.text_frame.paragraphs[0]; hp.alignment = PP_ALIGN.CENTER
            # 인간 _final 정합: 제목 자간 좁힘(-30) + 흰 윤곽선(가독성)
            _add_run(hp, head_txt, fsize=title_size, bold=True, color=WHITE,
                     spc=-30, outline={"color": "FFFFFF", "width_pt": 0.5})
            set_autofit_shrink(hb.text_frame)
            body_y = y + header_h + pad * 0.6
            body_h = h - header_h - pad * 1.6
        else:
            body_y = y + pad
            body_h = h - pad * 2

        # ③ 본문 (bullets 리스트 or 단락) — 동적 폰트 리사이징 + normAutofit
        bullets = item.get("bullets")
        body_w = card_w - 2 * pad
        if bullets:
            para_texts = ["• " + b for b in bullets]
        else:
            para_texts = [item.get("body", "")]
        # 파이썬 측 사전 리사이징(추정 줄수 기반) → 넘침 방지
        eff_size = fit_font_size(para_texts, body_w, body_h, body_size,
                                 min_pt=7.5)
        tb = slide.shapes.add_textbox(Inches(cx + pad), Inches(body_y),
                                      Inches(body_w), Inches(body_h))
        _set_no_fill(tb)
        tf = tb.text_frame; tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.04)
        tf.margin_top = tf.margin_bottom = Inches(0.03)
        for j, t in enumerate(para_texts):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            p.alignment = body_align
            _add_run(p, t, fsize=eff_size, color=DARK_GRAY, spc=-20)
            _set_line_spacing(p, 116)
        set_autofit_shrink(tf)   # PowerPoint 자동 축소 안전망
    return shapes


# ── 컨테이너(박스 안 박스) 레이아웃 (인간 _final slide 3·31 재현) ──
def add_container_box(slide, x, y, w, h, inner_items, fill=None,
                      title=None, title_size=12, inner_size=9.5,
                      inner_gap=0.20, label_h=0.0):
    """외곽 컨테이너 박스 + 내부 N개 텍스트 블록 (가로 분할) 배치.

    inner_items : [{"head":..,"body":..}, ...] 또는 문자열 리스트.
    fill        : 외곽 박스 채움색 (기본 CONTAINER_GR 회색).
    title       : 컨테이너 상단 라벨(옵션).
    인간본: 외곽 SOLID 박스(E6E6E6/5D8ECF) 안에 TextBox 3~4개 균등 배치.
    """
    fill = to_rgb(fill) if fill is not None else THEME["container"]
    outer = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
                                   Inches(x), Inches(y), Inches(w), Inches(h))
    outer.fill.solid(); outer.fill.fore_color.rgb = fill
    outer.line.fill.background()
    add_shadow(outer, blur_pt=5, dist_pt=2.5, alpha_pct=72)

    # 상대휘도로 어두운 배경 판정 → 텍스트 흰색 자동 전환
    dark_bg = _dark_bg(fill)
    head_color = WHITE if dark_bg else THEME["accent"]
    body_color = WHITE if dark_bg else DARK_GRAY
    top = y + 0.12
    if title:
        add_textbox(slide, title, x + 0.15, y + 0.06, w - 0.30, 0.34,
                    fsize=title_size, bold=True,
                    color=WHITE if dark_bg else DARK_GRAY)
        top = y + 0.48
    n = max(1, len(inner_items))
    pad = 0.18
    inner_w = (w - pad * 2 - inner_gap * (n - 1)) / n
    inner_h = h - (top - y) - 0.18
    for i, it in enumerate(inner_items):
        if isinstance(it, str):
            it = {"head": "", "body": it}
        ix = x + pad + i * (inner_w + inner_gap)
        # 동적 리사이징(head 1줄 + body) 추정
        ptexts = ([it["head"]] if it.get("head") else []) + [it.get("body", "")]
        eff = fit_font_size(ptexts, inner_w, inner_h, inner_size, min_pt=7.5)
        tb = slide.shapes.add_textbox(Inches(ix), Inches(top),
                                      Inches(inner_w), Inches(inner_h))
        tf = tb.text_frame; tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.04)
        tf.margin_top = tf.margin_bottom = Inches(0.03)
        p0 = tf.paragraphs[0]
        if it.get("head"):
            p0.alignment = PP_ALIGN.LEFT
            _add_run(p0, it["head"], fsize=eff + 1, bold=True,
                     color=head_color, spc=-20)
            pb = tf.add_paragraph()
        else:
            pb = p0
        pb.alignment = PP_ALIGN.LEFT
        _add_run(pb, it.get("body", ""), fsize=eff, color=body_color, spc=-20)
        _set_line_spacing(pb, 116)
        set_autofit_shrink(tf)
    return outer


# ════════════════════════════════════════════════════════════
# 통합 마스터 엔진 — DeckEngine (v5.0)
#   "어떤 프로젝트든 auto_ppt.py 하나만 호출 → 외부 데이터(spec)로 PPT 생성"
#   - 템플릿 로드 + 슬라이드 삭제 자동
#   - cover/toc/overview 일반화 (kia·lotte 중복 3종 통합)
#   - render(spec): 슬라이드 타입별 디스패치
#   사용: DeckEngine(template, out).render(spec).save()
#   CLI : python auto_ppt.py <spec.json>
# ════════════════════════════════════════════════════════════

class DeckEngine:
    SLIDE_W = 10.835   # 27.52cm — KIA/HCG 공통 마스터 실측
    SLIDE_H = 7.5      # 19.05cm

    def __init__(self, template=None, out=None, theme="hcg", design_tokens=None, color_overrides=None):
        # v6.0: best-practice 베이스 = 기아 _final. 없으면 롯데(REAL) 폴백.
        if template is None:
            template = KIA_FINAL if KIA_FINAL and os.path.exists(KIA_FINAL) else REAL
        self.template = template
        self.out = out or OUT
        # v7.0: 프로젝트 브랜드 테마 적용 (hcg/paradise). 도식 함수가 THEME 참조.
        self.theme = apply_theme(theme)
        if design_tokens or color_overrides:
            apply_design_tokens(design_tokens or {}, color_overrides)
        print(f"Loading base template (theme={theme}):", self.template)
        self.prs = Presentation(self.template)
        # 베이스 템플릿을 '깡통'이 아닌 디자인 상속용으로 사용 →
        # 슬라이드만 비우고 master/layout(표지·목차·본문)·테마·폰트는 보존.
        delete_all_slides(self.prs)
        self.count = 0
        # 레이아웃 이름→인덱스 매핑 (이름 기반 안전 접근)
        self.layout_by_name = {}
        for i, lay in enumerate(self.prs.slide_layouts):
            self.layout_by_name[lay.name] = i

    def _layout(self, name, fallback):
        """레이아웃을 이름으로 우선 탐색, 없으면 fallback 인덱스."""
        return self.layout_by_name.get(name, fallback)

    def _mark(self, label):
        self.count += 1
        print(f"  S{self.count:02d} {label}")
        return self.count

    # ── 일반화된 cover (kia/lotte 중복 통합) ──────────────────
    def cover(self, title, subtitle="- 제안서 -", date=None):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[0])
        for ph in slide.placeholders:
            idx = ph.placeholder_format.idx
            if idx == 0:
                ph.text = title
            elif idx == 1:
                ph.text = subtitle
        if date:
            add_textbox(slide, date, 1.137, 5.319, 8.563, 0.30,
                        fsize=10, align=PP_ALIGN.CENTER, color=DARK_GRAY)
        return slide

    # ── 목차/섹션 어젠다 (v6.0: 텍스트 나열 폐기 → 인간 카드 재현) ──
    def toc(self, items, title=None, current=None):
        """목차 슬라이드. items = ["섹션명", ...] 또는 [(번호,섹션명,페이지)..].
        구버전 튜플 입력도 호환 (섹션명만 추출). current=강조 인덱스.
        실제 도형 배치는 create_toc_slide()에 위임 (인간 _final 재현)."""
        sections = []
        for it in items:
            if isinstance(it, (list, tuple)):
                # (번호, 섹션명, 페이지) → 섹션명만 사용
                sections.append(str(it[1]) if len(it) > 1 else str(it[0]))
            else:
                sections.append(str(it))
        return create_toc_slide(self.prs, sections, current=current,
                                layout_idx=self._layout('목차', 1), title=title)

    def section_divider(self, sections, current):
        """섹션 디바이더 = 같은 어젠다카드에서 current 섹션만 강조 이동."""
        return create_toc_slide(self.prs, sections, current=current,
                                layout_idx=self._layout('목차', 1))

    # ── 구조화 콘텐츠 블록 슬라이드 (Text-to-Visual 핵심) ────────
    def content_blocks(self, title, subtitle, items, y=2.05, h=4.30,
                       accent=None, numbered=True, bar_label=None,
                       quote=None, example=False):
        """리스트형 텍스트 → N등분 카드 그리드 본문 슬라이드.
        items = [{"title","body"|"bullets"}, ...]."""
        slide = self.prs.slides.add_slide(
            self.prs.slide_layouts[self._layout('본문', 2)])
        set_title(slide, title, subtitle)
        if bar_label:
            add_header_bar(slide, bar_label)
            y = max(y, 2.05)
        add_structured_content_blocks(slide, items, y=y, h=h,
                                      accent=accent, numbered=numbered)
        if quote:
            add_insight_quote(slide, quote, x=CM(1.46), y=6.70, w=CM(24.71))
        if example:
            add_example_badge(slide)
        return slide

    # ── 컨테이너(박스 안 박스) 슬라이드 ──────────────────────────
    def container_slide(self, title, subtitle, containers):
        """containers = [{"x","y","w","h","fill","title","items":[...]}...]."""
        slide = self.prs.slides.add_slide(
            self.prs.slide_layouts[self._layout('본문', 2)])
        set_title(slide, title, subtitle)
        for c in containers:
            add_container_box(slide, c["x"], c["y"], c["w"], c["h"],
                              c["items"], fill=c.get("fill"),
                              title=c.get("title"))
        return slide

    # ── 일반화된 Project Overview (kia/lotte 중복 통합) ───────
    def project_overview(self, subtitle, background=None, scope=None,
                         sub_details=None, plan=None, quote=None,
                         bg_label="추진 배경", scope_label="추진 범위",
                         plan_label="추진 방안"):
        """scope = [(라벨, x), ...] 4단계 / sub_details = [(x, 설명), ...]
        plan = 문자열(여러 줄) / background = 문자열 / quote = takeaway."""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[2])
        set_title(slide, "Project Overview", subtitle)
        y = 1.66
        if background:
            add_textbox(slide, bg_label, LABEL_X, y, LABEL_W, 0.32,
                        fsize=12, italic=True, align=PP_ALIGN.RIGHT, color=MED_GRAY)
            add_textbox(slide, background, 1.836, y, 8.759, 0.55,
                        fsize=11, color=DARK_GRAY)
        if scope:
            add_textbox(slide, scope_label, 0.0, 2.58, 1.62, 0.32,
                        fsize=12, italic=True, align=PP_ALIGN.RIGHT, color=MED_GRAY)
            for t, x in scope:
                add_item(slide, t, x, 2.58, width=1.976, height=0.62,
                         fsize=11, align=PP_ALIGN.CENTER, fill_rgb=HCG_RED,
                         color=WHITE, no_border=True, shadow=True)
            centers = [(x + 0.99, 2.89) for _, x in scope]
            add_arrow_flow(slide, centers)
        if sub_details:
            for x, txt in sub_details:
                add_textbox(slide, txt, x, 3.28, 1.951, 0.55, fsize=9, color=DARK_GRAY)
        if plan:
            add_textbox(slide, plan_label, LABEL_X, 4.30, LABEL_W, 0.50,
                        fsize=12, italic=True, align=PP_ALIGN.RIGHT, color=MED_GRAY)
            add_textbox(slide, plan, 1.836, 4.30, 8.759, 1.35, fsize=10, color=DARK_GRAY)
        if quote:
            add_insight_quote(slide, quote, x=0.691, y=6.55, w=9.451)
        return slide

    # ── 고급 XML 헬퍼 데모 (Phase 1 갭 보강 시연) ────────────
    def demo_advanced(self, title="고급 XML 디자인 요소", subtitle=None):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[2])
        set_title(slide, title, subtitle or "윤곽선·그림자·그라디언트·투명도·자간 (인간 _final 정합)")
        # 그라디언트 박스 + 텍스트 윤곽선
        add_item(slide, "Gradient + Outline", 0.69, 2.2, width=2.9, height=1.1,
                 fsize=14, bold=True, align=PP_ALIGN.CENTER, color=WHITE,
                 no_border=True, grad=(HCG_RED, CORAL_DK, 45),
                 outline={"color": "FFFFFF", "width_pt": 1.0}, shadow=True)
        # 그림자 박스
        add_item(slide, "Drop Shadow", 3.95, 2.2, width=2.9, height=1.1,
                 fsize=14, bold=True, align=PP_ALIGN.CENTER, color=WHITE,
                 no_border=True, fill_rgb=BLUE, shadow=True)
        # 반투명 박스 (겹침 표현)
        base = add_item(slide, "", 7.2, 2.2, width=2.9, height=1.1,
                        fill_rgb=DARK_GRAY, no_border=True)
        over = add_item(slide, "Transparency 40%", 7.5, 2.5, width=2.4, height=0.9,
                        fsize=12, bold=True, align=PP_ALIGN.CENTER, color=WHITE,
                        no_border=True, fill_rgb=HCG_RED)
        set_transparency(over, 40)
        # 자간(tight) 시연
        add_item(slide, "자간 -50 적용 타이틀", 0.69, 3.7, width=9.45, height=0.6,
                 fsize=16, bold=True, align=PP_ALIGN.CENTER, color=HCG_RED,
                 no_border=True, fill_rgb=GRAY_LT, spc=-50, anchor='ctr')
        add_insight_quote(slide, "XML 레벨 제어로 인간 제작본 디자인 밀도에 근접", 0.69, 6.6, 9.45)
        return slide

    # ── 데이터 구동 렌더러 ────────────────────────────────────
    def render(self, spec):
        """spec = {"meta": {...}, "slides": [ {type, ...}, ... ]}"""
        prs = self.prs
        for s in spec.get("slides", []):
            t = s.get("type")
            if t == "cover":
                self.cover(s["title"], s.get("subtitle", "- 제안서 -"), s.get("date"))
            elif t == "toc":
                # v6.0: 섹션명 리스트(권장) 또는 (번호,섹션명,페이지) 튜플 호환
                self.toc(s["items"], s.get("title"), s.get("current"))
            elif t == "section_agenda":
                self.section_divider(s["sections"], s["current"])
            elif t == "blocks":
                self.content_blocks(s["title"], s["subtitle"], s["items"],
                                    y=s.get("y", 2.05), h=s.get("h", 4.30),
                                    numbered=s.get("numbered", True),
                                    bar_label=s.get("bar_label"),
                                    quote=s.get("quote"),
                                    example=s.get("example", False))
            elif t == "container":
                self.container_slide(s["title"], s["subtitle"], s["containers"])
            elif t == "overview":
                self.project_overview(
                    s["subtitle"], s.get("background"),
                    [tuple(x) for x in s.get("scope", [])] or None,
                    [tuple(x) for x in s.get("sub_details", [])] or None,
                    s.get("plan"), s.get("quote"))
            elif t == "section":
                build_section_divider(prs, s["num"], s["title"], s.get("sub", ""))
            elif t == "body_2col":
                build_body_2col(prs, s["title"], s["subtitle"], s["header_l"],
                                s["header_r"], s["left"], s["right"])
            elif t == "body_single":
                build_body_single(prs, s["title"], s["subtitle"],
                                  [tuple(r) for r in s["rows"]])
            elif t == "body_process":
                build_body_process(prs, s["title"], s["subtitle"],
                                   [tuple(x) for x in s["steps"]],
                                   [tuple(x) for x in s.get("desc", [])] or None)
            elif t == "overview_3col":
                build_overview_3col(prs, s["title"], s["subtitle"], s["cols"],
                                    s.get("bar_label"), s.get("example", False))
            elif t == "diff_matrix":
                build_diff_matrix(prs, s["title"], s["subtitle"], s["rows"],
                                  s.get("quote"), s.get("example", False))
            elif t == "pain_point_categorized":
                build_pain_point_categorized(prs, s["title"], s["subtitle"],
                                             s["left_hdr"], s["right_hdr"],
                                             [tuple(r) for r in s["rows"]],
                                             s.get("summary_left"), s.get("summary_right"))
            elif t == "approach_vs":
                build_approach_vs(prs, s["title"], s["subtitle"], s["left_title"],
                                  s["left"], s["right_title"], s["right"], s.get("quote"))
            elif t == "process_roadmap":
                build_process_roadmap(prs, s["title"], s["subtitle"], s["phases"])
            elif t == "compare_table":
                build_compare_table(prs, s["title"], s["subtitle"], s["headers"],
                                    s["rows"], s.get("example", False))
            elif t == "appendix":
                build_body_single(prs, s.get("title", "Appendix"), s.get("subtitle", ""),
                                  [tuple(r) for r in s["rows"]])
            elif t == "demo_advanced":
                self.demo_advanced(s.get("title", "고급 XML 디자인 요소"), s.get("subtitle"))
            elif t == "end":
                build_end(prs)
            else:
                print("  [skip] unknown slide type:", t)
                continue
            self._mark(s.get("label", t))
        return self

    def save(self, out=None):
        path = out or self.out
        print(f"\nSaving → {path}")
        self.prs.save(path)
        print("Done!")
        return path

def build_showcase(out=None):
    """v6.0 쇼케이스 — 기아 _final 템플릿 위에 신규 엔진 기능 시연.
    목차(어젠다카드) + 섹션 디바이더 + 구조화 카드 그리드(3/4분할) + 컨테이너."""
    out = out or "C:\\Users\\cgpar\\ppt-skill\\HCG_Automated_Draft.pptx"
    eng = DeckEngine(template=None, out=out)   # None → KIA_FINAL 자동
    sections = ["제안 배경", "수행 방안", "일정 및 조직"]

    # S1 표지
    eng.cover("기아 중장기 보상체계 개선 추진", "- 제안서 -", "2025. 09")
    eng._mark("Cover")

    # S2 목차 (전체 어젠다 — 첫 섹션 강조)
    eng.toc(sections, current=0)
    eng._mark("TOC agenda")

    # S3 섹션 디바이더 (II 강조로 이동 — 동일 카드, 하이라이트만 하강)
    eng.section_divider(sections, current=1)
    eng._mark("Section II divider")

    # S4 구조화 3-카드 그리드 (Text-to-Visual)
    eng.content_blocks(
        "Why HCG", "HCG는 3대 핵심 역량을 통해 금번 프로젝트를 성공적으로 추진",
        [{"title": "자동차 산업 이해",
          "bullets": ["현대/기아 그룹 계열사 다수 수행", "산업 보상 트렌드 보유"]},
         {"title": "대기업 보상 경험",
          "bullets": ["국내 대기업 보상 프로젝트 다수", "복잡·세분화 주제 전문성"]},
         {"title": "직군별 차별화 노하우",
          "bullets": ["직군 특화 방법론 구축", "Key Question 기반 추진"]}],
        quote="국내 최고 인사/조직 전문 컨설팅 역량으로 성공적 추진")
    eng._mark("3-card blocks")

    # S5 구조화 4-카드 그리드 (번호형, 자동 폰트 축소)
    eng.content_blocks(
        "수행 방안 Overview", "4단계 접근으로 중장기 보상체계 개선 추진",
        [{"title": "벤치마킹", "body": "글로벌/로컬 보상 수준 진단 및 시장 경쟁력 분석"},
         {"title": "현황 진단", "body": "내부 보상 데이터 분석 및 Pain Point 도출"},
         {"title": "체계 설계", "body": "직무가치 기반 Pay Band 및 성과연동 보상 설계"},
         {"title": "실행 로드맵", "body": "단계적 전환 계획 및 Change Management"}],
        bar_label="중장기 보상체계 개선 4-Step Approach")
    eng._mark("4-card blocks")

    # S6 컨테이너(박스 안 박스)
    eng.container_slide(
        "Global Big Data Analytics", "외부 플랫폼 + 자체 데이터 융합 분석 체계",
        [{"x": CM(4.86), "y": 2.10, "w": CM(20.89), "h": 1.95,
          "fill": CONTAINER_GR, "title": None,
          "items": [{"head": "활성 이용자 데이터", "body": "월 1,000만명 제출 보상 데이터"},
                    {"head": "자체 플랫폼", "body": "실시간 데이터 보유"},
                    {"head": "서베이 연계", "body": "Mercer/Radford 등 확보"}]}])
    eng._mark("Container box-in-box")

    eng.save()
    return out


# --- Unified-framework additions (additive; default behavior unchanged) ---

Designer = DeckEngine  # public alias for the framework


def apply_design_tokens(tokens, overrides=None):
    """Update the module color palette from parsed skill_ppt_design.json.

    `tokens` is the parsed design JSON; reads tokens["theme_colors"], where each
    value is either a hex string or a dict with an "rgb" key. `overrides` is an
    optional {NAME: "#RRGGBB"} map (e.g. client identity.colors) applied last.
    No-op for names not present. Mirrors apply_theme()'s global-rebind pattern.
    """
    global HCG_RED
    resolved = {}
    for name, val in (tokens.get("theme_colors") or {}).items():
        hexv = val.get("rgb") if isinstance(val, dict) else val
        if hexv:
            resolved[name] = str(hexv).lstrip("#")
    for name, hexv in (overrides or {}).items():
        if hexv:
            resolved[name] = str(hexv).lstrip("#")
    if "HCG_RED" in resolved:
        HCG_RED = RGBColor.from_string(resolved["HCG_RED"])
    for key, hexv in resolved.items():
        if isinstance(THEME, dict) and key in THEME:
            THEME[key] = RGBColor.from_string(hexv)
