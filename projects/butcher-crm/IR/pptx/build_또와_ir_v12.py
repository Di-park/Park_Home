#!/usr/bin/env python3
"""
또와 IR 덱 v1.2 — PowerPoint builder (AWS tone, 16 slides)

Aligns with: projects/butcher-crm/IR/또와_IR덱_v0.8.md

Output: /tmp/butcher-build/또와_IR덱_v1.2_AWS.pptx

Anti-breakage 5 principles:
  1. Font: <a:latin>, <a:ea>, <a:cs> 3 slots all forced to Noto Sans KR
  2. No emojis (replaced with [O]/[!]/[X] text or basic shapes)
  3. word_wrap=True + boxes sized for content + line_spacing explicit
  4. Simple boxes + straight connectors only
  5. Manual boxes (no native tables)
"""

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ============================================================
# Design system — AWS tone
# ============================================================

# AWS Orange palette
AWS_ORANGE       = RGBColor(0xD1, 0x51, 0x27)  # #D15127 main
AWS_ORANGE_DARK  = RGBColor(0xA8, 0x40, 0x1E)  # #A8401E
AWS_ORANGE_LIGHT = RGBColor(0xFA, 0xD9, 0xCB)  # #FAD9CB
AWS_ORANGE_FAINT = RGBColor(0xFF, 0xF7, 0xF2)  # #FFF7F2 (card bg)
AWS_ORANGE_MID   = RGBColor(0xE8, 0x94, 0x76)  # #E89476
AWS_DARK_NAVY    = RGBColor(0x23, 0x2F, 0x3E)  # #232F3E

# Grayscale
GRAY_900 = RGBColor(0x1A, 0x1A, 0x1A)
GRAY_800 = RGBColor(0x33, 0x33, 0x33)
GRAY_700 = RGBColor(0x4D, 0x4D, 0x4D)
GRAY_600 = RGBColor(0x66, 0x66, 0x66)
GRAY_500 = RGBColor(0x80, 0x80, 0x80)
GRAY_400 = RGBColor(0xB3, 0xB3, 0xB3)
GRAY_300 = RGBColor(0xCC, 0xCC, 0xCC)
GRAY_200 = RGBColor(0xE6, 0xE6, 0xE6)
GRAY_100 = RGBColor(0xF2, 0xF2, 0xF2)
GRAY_050 = RGBColor(0xFA, 0xFA, 0xFA)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)

# Accent colors
GREEN_OK  = RGBColor(0x16, 0xA3, 0x4A)  # ✓
AMBER     = RGBColor(0xF5, 0x9E, 0x0B)  # !
RED       = RGBColor(0xDC, 0x26, 0x26)  # ✗
PURPLE    = RGBColor(0x7C, 0x3A, 0xED)
BROWN     = RGBColor(0x8B, 0x5C, 0x32)  # meat-ish

# Font
FONT_KR = "Noto Sans KR"

# Slide geometry (16:9)
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# Margins
M_LEFT   = Inches(0.6)
M_RIGHT  = Inches(0.6)
M_TOP    = Inches(0.5)
M_BOTTOM = Inches(0.4)

CONTENT_W = SLIDE_W - M_LEFT - M_RIGHT  # 12.133"
TITLE_TOP = Inches(0.5)
TITLE_H   = Inches(0.9)
BODY_TOP  = Inches(1.5)
BODY_H    = Inches(5.4)
PAGE_NUM_TOP = Inches(7.05)


# ============================================================
# XML font helpers (3-slot enforcement)
# ============================================================

def _force_font_xml(run):
    """Force Noto Sans KR in <a:latin>, <a:ea>, <a:cs> 3 slots."""
    rPr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        # Remove existing
        for el in rPr.findall(qn(f"a:{tag}")):
            rPr.remove(el)
        # Add fresh
        el = etree.SubElement(rPr, qn(f"a:{tag}"))
        el.set("typeface", FONT_KR)


def set_run(run, text, *, size=14, bold=False, italic=False, color=GRAY_800, font=FONT_KR):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    _force_font_xml(run)


# ============================================================
# Shape helpers
# ============================================================

def add_rect(slide, left, top, width, height, *, fill=WHITE, line=None, line_w=0.75, corner=False, shadow=False):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if corner else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_w)
    if not shadow:
        # disable shadow
        sp = shp.shadow
        try:
            sp.inherit = False
        except Exception:
            pass
    if corner:
        # Set corner radius small (default is large)
        try:
            shp.adjustments[0] = 0.08
        except Exception:
            pass
    return shp


def add_text(slide, left, top, width, height, *,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, autosize=False, word_wrap=True):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = word_wrap
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.03)
    tf.vertical_anchor = anchor
    # First paragraph alignment default
    tf.paragraphs[0].alignment = align
    return tb, tf


def add_para(tf, text="", *, size=14, bold=False, italic=False, color=GRAY_800,
             align=PP_ALIGN.LEFT, line_spacing=1.25, space_before=0, space_after=0, first=False):
    if first:
        p = tf.paragraphs[0]
        # clear default run
        for r in list(p.runs):
            r.text = ""
    else:
        p = tf.add_paragraph()
    p.alignment = align
    p.line_spacing = line_spacing
    if space_before:
        p.space_before = Pt(space_before)
    if space_after:
        p.space_after = Pt(space_after)
    run = p.add_run()
    set_run(run, text, size=size, bold=bold, italic=italic, color=color)
    return p, run


def add_line(slide, x1, y1, x2, y2, *, color=GRAY_300, weight=0.75):
    conn = slide.shapes.add_connector(1, x1, y1, x2 - x1, y2 - y1)  # 1=STRAIGHT
    conn.line.color.rgb = color
    conn.line.width = Pt(weight)
    return conn


# ============================================================
# Slide chrome (title, page number, accent bar)
# ============================================================

def add_chrome(slide, page_num, total, title_text, *, eyebrow=None, subtitle=None):
    # Top-left accent bar
    add_rect(slide, M_LEFT, TITLE_TOP, Inches(0.08), Inches(0.4),
             fill=AWS_ORANGE, line=None)

    # Eyebrow (small text above title)
    title_y = TITLE_TOP
    if eyebrow:
        _, tf = add_text(slide, M_LEFT + Inches(0.2), TITLE_TOP - Inches(0.05),
                          CONTENT_W, Inches(0.3))
        add_para(tf, eyebrow, size=11, bold=True, color=AWS_ORANGE,
                 line_spacing=1.0, first=True)
        title_y = TITLE_TOP + Inches(0.25)

    # Title
    _, tf = add_text(slide, M_LEFT + Inches(0.2), title_y,
                     CONTENT_W - Inches(0.2), Inches(0.6))
    add_para(tf, title_text, size=24, bold=True, color=GRAY_900,
             line_spacing=1.15, first=True)

    # Subtitle
    if subtitle:
        _, tf = add_text(slide, M_LEFT + Inches(0.2), title_y + Inches(0.6),
                         CONTENT_W, Inches(0.4))
        add_para(tf, subtitle, size=13, italic=True, color=GRAY_600,
                 line_spacing=1.2, first=True)

    # Page number (bottom-right)
    _, tf = add_text(slide, SLIDE_W - Inches(1.6), PAGE_NUM_TOP,
                     Inches(1.3), Inches(0.3),
                     align=PP_ALIGN.RIGHT)
    add_para(tf, f"{page_num} / {total}", size=10, color=GRAY_400,
             line_spacing=1.0, align=PP_ALIGN.RIGHT, first=True)


# ============================================================
# Slide builders — 16 slides
# ============================================================

TOTAL = 16


def slide_blank(prs):
    layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(layout)
    # Force white bg (slide background; we draw white rect to be safe)
    bg = add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill=WHITE, line=None)
    bg.shadow.inherit = False
    return slide


# ---------- Slide 1: Cover ----------
def s01_cover(prs):
    slide = slide_blank(prs)

    # Left orange band
    add_rect(slide, 0, 0, Inches(0.3), SLIDE_H, fill=AWS_ORANGE, line=None)

    # Company tag
    _, tf = add_text(slide, Inches(1.0), Inches(1.4), Inches(10), Inches(0.4))
    add_para(tf, "㈜위브원  |  또와", size=14, bold=True, color=AWS_ORANGE,
             line_spacing=1.0, first=True)

    # Main title
    _, tf = add_text(slide, Inches(1.0), Inches(2.1), Inches(11.5), Inches(2.0))
    add_para(tf, "정육점 특화", size=42, bold=False, color=GRAY_900,
             line_spacing=1.1, first=True)
    add_para(tf, "Zero-Task AI 에이전트", size=52, bold=True, color=AWS_ORANGE_DARK,
             line_spacing=1.1)

    # Tagline
    _, tf = add_text(slide, Inches(1.0), Inches(4.5), Inches(11.5), Inches(0.6))
    add_para(tf, "단골부터 마진까지, 알아서 다 챙기는 막내 직원.",
             size=20, italic=True, color=GRAY_700, line_spacing=1.2, first=True)

    # Speaker block (bottom)
    add_line(slide, Inches(1.0), Inches(6.0), Inches(6.0), Inches(6.0),
             color=AWS_ORANGE, weight=1.5)
    _, tf = add_text(slide, Inches(1.0), Inches(6.1), Inches(8), Inches(0.6))
    add_para(tf, "발표: 안동철  ·  대표이사  ·  2026.05",
             size=13, color=GRAY_600, line_spacing=1.2, first=True)

    # Page indicator (cover-style)
    _, tf = add_text(slide, SLIDE_W - Inches(1.6), PAGE_NUM_TOP,
                     Inches(1.3), Inches(0.3), align=PP_ALIGN.RIGHT)
    add_para(tf, f"1 / {TOTAL}", size=10, color=GRAY_400,
             line_spacing=1.0, align=PP_ALIGN.RIGHT, first=True)


# ---------- Slide 2: One-Liner ----------
def s02_oneliner(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 2, TOTAL,
               "또와는 고객관리부터 마케팅·재고 관리·발주까지 스스로 하는,",
               eyebrow="ONE-LINER",
               subtitle="정육점 사장님의 첫 AI 직원입니다.")

    # 3-card Zero-Task promise
    card_y = Inches(2.8)
    card_h = Inches(2.0)
    gap = Inches(0.3)
    card_w = (CONTENT_W - gap * 2) / 3

    items = [
        ("0분", "단골 챙기는 시간", AWS_ORANGE),
        ("0분", "메시지 쓰는 시간", AWS_ORANGE_DARK),
        ("주 5초", "사장님이 남기는 시간\n(리포트 확인)", AWS_DARK_NAVY),
    ]
    for i, (big, small, color) in enumerate(items):
        x = M_LEFT + (card_w + gap) * i
        # Card bg
        add_rect(slide, x, card_y, card_w, card_h,
                 fill=AWS_ORANGE_FAINT, line=AWS_ORANGE_LIGHT, line_w=1.0, corner=True)
        # Big number
        _, tf = add_text(slide, x, card_y + Inches(0.3), card_w, Inches(0.9),
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, big, size=44, bold=True, color=color,
                 align=PP_ALIGN.CENTER, line_spacing=1.0, first=True)
        # Small label
        _, tf = add_text(slide, x, card_y + Inches(1.3), card_w, Inches(0.6),
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, small, size=14, color=GRAY_700,
                 align=PP_ALIGN.CENTER, line_spacing=1.3, first=True)

    # Bottom caption
    _, tf = add_text(slide, M_LEFT, Inches(5.4), CONTENT_W, Inches(0.5),
                     align=PP_ALIGN.CENTER)
    add_para(tf,
             "단골관리에서 시작해 마진·발주까지 — 또와가 매장 운영을 알아서 챙깁니다.",
             size=14, italic=True, color=GRAY_600,
             align=PP_ALIGN.CENTER, line_spacing=1.3, first=True)


# ---------- Slide 3: Problem ----------
def s03_problem(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 3, TOTAL,
               "정육점 매출의 70%는 단골에서. 그런데 사장님은 단골을 잃는 줄도 모릅니다.",
               eyebrow="PROBLEM")

    # 3 problem cards (horizontal)
    card_y = Inches(2.6)
    card_h = Inches(2.7)
    gap = Inches(0.25)
    card_w = (CONTENT_W - gap * 2) / 3

    items = [
        ("01", "단골 정보가 사장님 머릿속에만 있다",
         "파일럿 1매장만 봐도 단골 약 2,100명.\n사장님이 이름·패턴을 인지하는 단골은 극소수."),
        ("02", "단골이 떠나기 전 어떠한 신호도 알 수 없다",
         "방문 주기가 늘어나도, 객단가가 떨어져도 사장님은 모른다.\n다른 가게로 옮긴 후에야 알아챈다."),
        ("03", "단골을 관리할 수 있는 여유가 없다",
         "아침부터 저녁까지 장사 준비·손질·응대·마감까지 다 해야 하는 사장님에겐, 체계적으로 고객을 관리할 시간적·비용적 여유가 부족."),
    ]
    for i, (num, title, body) in enumerate(items):
        x = M_LEFT + (card_w + gap) * i
        add_rect(slide, x, card_y, card_w, card_h,
                 fill=WHITE, line=GRAY_300, line_w=1.0, corner=True)
        # Number badge
        add_rect(slide, x + Inches(0.3), card_y + Inches(0.3),
                 Inches(0.5), Inches(0.5),
                 fill=AWS_ORANGE, line=None, corner=True)
        _, tf = add_text(slide, x + Inches(0.3), card_y + Inches(0.3),
                          Inches(0.5), Inches(0.5),
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, num, size=14, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, line_spacing=1.0, first=True)
        # Title
        _, tf = add_text(slide, x + Inches(0.3), card_y + Inches(0.95),
                          card_w - Inches(0.6), Inches(0.7))
        add_para(tf, title, size=15, bold=True, color=GRAY_900,
                 line_spacing=1.3, first=True)
        # Body
        _, tf = add_text(slide, x + Inches(0.3), card_y + Inches(1.7),
                          card_w - Inches(0.6), card_h - Inches(1.9))
        add_para(tf, body, size=12, color=GRAY_700, line_spacing=1.4, first=True)


# ---------- Slide 4: Solution ----------
def s04_solution(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 4, TOTAL,
               "사장님은 눈앞의 손님에게 집중하세요.",
               eyebrow="SOLUTION",
               subtitle="나머지는 또와가 알아서 합니다.")

    # Left: 3 metric cards (vertical stack)
    left_x = M_LEFT
    left_w = Inches(4.5)
    metric_top = Inches(2.7)
    metric_h = Inches(1.0)
    metric_gap = Inches(0.15)

    _, tf = add_text(slide, left_x, Inches(2.3), left_w, Inches(0.4))
    add_para(tf, "효과 (제품 가치 가설)", size=12, bold=True,
             color=AWS_ORANGE_DARK, line_spacing=1.0, first=True)

    metrics = [
        ("주 5초", "시간 절약 — 사장님이 단골관리에 쓰는 시간"),
        ("−50%", "비용 감소 — 마케팅·재고 비용"),
        ("+15%", "매출 증가 — 매장 매출"),
    ]
    for i, (val, desc) in enumerate(metrics):
        y = metric_top + (metric_h + metric_gap) * i
        add_rect(slide, left_x, y, left_w, metric_h,
                 fill=AWS_ORANGE_FAINT, line=AWS_ORANGE_LIGHT, line_w=1.0, corner=True)
        _, tf = add_text(slide, left_x + Inches(0.3), y, Inches(1.5), metric_h,
                          anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, val, size=24, bold=True, color=AWS_ORANGE_DARK,
                 line_spacing=1.0, first=True)
        _, tf = add_text(slide, left_x + Inches(1.9), y,
                          left_w - Inches(2.1), metric_h, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, desc, size=12, color=GRAY_800, line_spacing=1.3, first=True)

    # Right: 3 function items
    right_x = M_LEFT + left_w + Inches(0.3)
    right_w = CONTENT_W - left_w - Inches(0.3)
    func_top = Inches(2.7)
    func_h = Inches(1.0)

    _, tf = add_text(slide, right_x, Inches(2.3), right_w, Inches(0.4))
    add_para(tf, "기능 — 3축", size=12, bold=True,
             color=AWS_ORANGE_DARK, line_spacing=1.0, first=True)

    funcs = [
        ("①", "매장 데이터 기반 단골 관리",
         "POS 거래 → 단골 자동 분류 · 메시지 발송 · 재방문 추적 (Wedge)"),
        ("②", "재고량 분석 CRM 활동",
         "과재고 부위 자동 마케팅 · 신선도 기반 추천"),
        ("③", "판매량 예측 기반 발주 제안·자동화",
         "단골 패턴 + 명절·계절 변수 → 발주량 자동"),
    ]
    for i, (num, title, body) in enumerate(funcs):
        y = func_top + (func_h + Inches(0.15)) * i
        add_rect(slide, right_x, y, right_w, func_h,
                 fill=WHITE, line=GRAY_300, line_w=1.0, corner=True)
        _, tf = add_text(slide, right_x + Inches(0.2), y, Inches(0.5), func_h,
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, num, size=22, bold=True, color=AWS_ORANGE,
                 align=PP_ALIGN.CENTER, line_spacing=1.0, first=True)
        _, tf = add_text(slide, right_x + Inches(0.7), y + Inches(0.1),
                          right_w - Inches(0.8), Inches(0.4))
        add_para(tf, title, size=13, bold=True, color=GRAY_900,
                 line_spacing=1.2, first=True)
        _, tf = add_text(slide, right_x + Inches(0.7), y + Inches(0.45),
                          right_w - Inches(0.8), Inches(0.5))
        add_para(tf, body, size=11, color=GRAY_700, line_spacing=1.3, first=True)

    # Bottom caption
    _, tf = add_text(slide, M_LEFT, Inches(6.4), CONTENT_W, Inches(0.4))
    add_para(tf, "* 효과 수치는 제품 가치 가설 — 실증 데이터 없음, 베타 30매장에서 측정 예정.",
             size=10, italic=True, color=GRAY_500, line_spacing=1.2, first=True)


# ---------- Slide 5: Product Concept Image (NEW) ----------
def s05_concept(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 5, TOTAL,
               "또와가 일하는 모습",
               eyebrow="PRODUCT CONCEPT",
               subtitle="사장님은 눈앞의 손님에게만 집중. 또와가 동시에 단골·재고·이탈을 챙깁니다.")

    # Big image placeholder area (left 55%)
    img_x = M_LEFT
    img_y = Inches(2.4)
    img_w = Inches(6.5)
    img_h = Inches(4.4)
    add_rect(slide, img_x, img_y, img_w, img_h,
             fill=AWS_ORANGE_FAINT, line=AWS_ORANGE_LIGHT, line_w=1.5, corner=True)
    _, tf = add_text(slide, img_x, img_y, img_w, img_h,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "[ 컨셉 이미지 영역 ]", size=16, bold=True, color=AWS_ORANGE_DARK,
             align=PP_ALIGN.CENTER, line_spacing=1.3, first=True)
    add_para(tf, "정육점 사장님이 손님 응대 중 +",
             size=12, color=GRAY_700, align=PP_ALIGN.CENTER, line_spacing=1.4)
    add_para(tf, "반투명 고스트 3마리가 동시에 일하는 컷",
             size=12, color=GRAY_700, align=PP_ALIGN.CENTER, line_spacing=1.4)
    add_para(tf, " ", size=10, color=GRAY_500,
             align=PP_ALIGN.CENTER, line_spacing=1.0)
    add_para(tf, "이미지 spec: assets/slide5_concept_image_prompt.md",
             size=9, italic=True, color=GRAY_500,
             align=PP_ALIGN.CENTER, line_spacing=1.2)

    # Right: 3 ghost bubble cards (vertical)
    bub_x = M_LEFT + img_w + Inches(0.3)
    bub_w = CONTENT_W - img_w - Inches(0.3)
    bub_top = Inches(2.4)
    bub_h = Inches(1.35)
    bub_gap = Inches(0.15)

    bubbles = [
        ("A", "정육왕 고객님! 보낸 문자 보고 찾아주셨군요! 감사해요!",
         "단골관리 · 재방문 유도"),
        ("B", "엇! 사장님, 목살 재고가 평소보다 많아요! 지금 프로모션 돌릴게요!",
         "재고 분석 · 자동 마케팅"),
        ("C", "사장님! 오늘 기준 이탈 고객 N명이 예상되는데요. 쿠폰 보내볼까요?",
         "이탈 예측 · 자동 대응"),
    ]
    for i, (label, msg, tag) in enumerate(bubbles):
        y = bub_top + (bub_h + bub_gap) * i
        # Bubble background
        add_rect(slide, bub_x, y, bub_w, bub_h,
                 fill=WHITE, line=AWS_ORANGE, line_w=1.5, corner=True)
        # Ghost label badge
        add_rect(slide, bub_x + Inches(0.2), y + Inches(0.2),
                 Inches(0.5), Inches(0.5),
                 fill=AWS_ORANGE, line=None, corner=True)
        _, tf = add_text(slide, bub_x + Inches(0.2), y + Inches(0.2),
                          Inches(0.5), Inches(0.5),
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, label, size=14, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, line_spacing=1.0, first=True)
        # Message
        _, tf = add_text(slide, bub_x + Inches(0.85), y + Inches(0.1),
                          bub_w - Inches(1.0), Inches(0.85))
        add_para(tf, f"“{msg}”", size=11, italic=True, color=GRAY_900,
                 line_spacing=1.3, first=True)
        # Tag
        _, tf = add_text(slide, bub_x + Inches(0.85), y + Inches(0.95),
                          bub_w - Inches(1.0), Inches(0.3))
        add_para(tf, f"→ {tag}", size=10, bold=True, color=AWS_ORANGE_DARK,
                 line_spacing=1.2, first=True)


# ---------- Slide 6: Why Now ----------
def s06_whynow(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 6, TOTAL,
               "도메인 특화 AI 에이전트의 시대. 소상공인 자리는 아직 비어있습니다.",
               eyebrow="WHY NOW")

    # 3 column cards
    col_y = Inches(2.4)
    col_h = Inches(3.4)
    gap = Inches(0.25)
    col_w = (CONTENT_W - gap * 2) / 3

    cols = [
        ("01", "AI 활용 비용이 1/10 이하로",
         "LLM API 비용이 3년 사이 폭락. 매장 월 몇 만원으로도 풀스택 AI 운영 가능한 첫 시점."),
        ("02", "도메인 특화 AI 에이전트가 등장 시작",
         "법률(Harvey), 의료(Hippocratic AI), 영업(11x) — 수직 도메인에서 한 업무를 완전 자동화하는 AI들이 글로벌 시장에 등장."),
        ("03", "그런데 소상공인 버티컬은 비어있음",
         "전문직·테크 도메인엔 도메인 AI가 빠르게 들어왔지만, 동네 자영업 도메인은 아직 0개사. 또와가 정육에서 시작."),
    ]
    for i, (num, title, body) in enumerate(cols):
        x = M_LEFT + (col_w + gap) * i
        add_rect(slide, x, col_y, col_w, col_h,
                 fill=WHITE, line=GRAY_300, line_w=1.0, corner=True)
        # Top accent bar
        add_rect(slide, x, col_y, col_w, Inches(0.08),
                 fill=AWS_ORANGE, line=None)
        # Number
        _, tf = add_text(slide, x + Inches(0.3), col_y + Inches(0.3),
                          col_w - Inches(0.6), Inches(0.5))
        add_para(tf, num, size=22, bold=True, color=AWS_ORANGE_DARK,
                 line_spacing=1.0, first=True)
        # Title
        _, tf = add_text(slide, x + Inches(0.3), col_y + Inches(0.95),
                          col_w - Inches(0.6), Inches(0.8))
        add_para(tf, title, size=15, bold=True, color=GRAY_900,
                 line_spacing=1.3, first=True)
        # Body
        _, tf = add_text(slide, x + Inches(0.3), col_y + Inches(1.85),
                          col_w - Inches(0.6), col_h - Inches(2.05))
        add_para(tf, body, size=11, color=GRAY_700, line_spacing=1.5, first=True)

    # Bottom callout strip
    callout_y = col_y + col_h + Inches(0.25)
    add_rect(slide, M_LEFT, callout_y, CONTENT_W, Inches(0.55),
             fill=AWS_ORANGE_FAINT, line=AWS_ORANGE_LIGHT, line_w=1.0, corner=True)
    _, tf = add_text(slide, M_LEFT, callout_y, CONTENT_W, Inches(0.55),
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf,
             "매크로 순풍 — 한국 1인당 정육 소비 54.6→62.9kg (5년 +15%)  ·  정부 소상공인 디지털 전환 바우처 연 수천억원",
             size=11, italic=True, color=GRAY_700,
             align=PP_ALIGN.CENTER, line_spacing=1.2, first=True)


# ---------- Slide 7: Market ----------
def s07_market(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 7, TOTAL,
               "정육 → 식품 소매·외식 → 재방문 B2C 소상공인 전체.",
               eyebrow="MARKET")

    # Top: TAM/SAM/SOM 3 stacked horizontally
    bar_y = Inches(2.3)
    bar_h = Inches(1.5)
    gap = Inches(0.3)
    bar_w = (CONTENT_W - gap * 2) / 3

    bars = [
        ("TAM", "재방문 기반 B2C 소상공인", "약 300만 업체 / 7조원", AWS_ORANGE_FAINT, AWS_ORANGE_LIGHT),
        ("SAM", "식품 관련 소매·외식", "약 90만 개 / 1조원", AWS_ORANGE_LIGHT, AWS_ORANGE_MID),
        ("SOM", "정육점 (첫 진입 시장)", "약 57,000개소 / 2,200억원", AWS_ORANGE, AWS_ORANGE_DARK),
    ]
    for i, (lbl, desc, size, fill, line) in enumerate(bars):
        x = M_LEFT + (bar_w + gap) * i
        text_color = WHITE if i == 2 else GRAY_900
        sub_color = WHITE if i == 2 else GRAY_700
        add_rect(slide, x, bar_y, bar_w, bar_h,
                 fill=fill, line=line, line_w=1.5, corner=True)
        _, tf = add_text(slide, x + Inches(0.3), bar_y + Inches(0.2),
                          bar_w - Inches(0.6), Inches(0.5))
        add_para(tf, lbl, size=20, bold=True, color=text_color,
                 line_spacing=1.0, first=True)
        _, tf = add_text(slide, x + Inches(0.3), bar_y + Inches(0.7),
                          bar_w - Inches(0.6), Inches(0.35))
        add_para(tf, desc, size=11, color=sub_color, line_spacing=1.2, first=True)
        _, tf = add_text(slide, x + Inches(0.3), bar_y + Inches(1.0),
                          bar_w - Inches(0.6), Inches(0.45))
        add_para(tf, size, size=14, bold=True, color=text_color,
                 line_spacing=1.2, first=True)

    # Bottom: Capturable SOM table
    tbl_y = Inches(4.2)
    tbl_h = Inches(2.2)
    tbl_w = CONTENT_W

    _, tf = add_text(slide, M_LEFT, Inches(4.0), tbl_w, Inches(0.3))
    add_para(tf, "Capturable SOM — 시간별 점유율 (벤치마크 기반)",
             size=12, bold=True, color=AWS_ORANGE_DARK,
             line_spacing=1.0, first=True)

    # Header
    cols_x = [M_LEFT,
              M_LEFT + Inches(2.5),
              M_LEFT + Inches(5.0),
              M_LEFT + Inches(8.0)]
    cols_w = [Inches(2.5), Inches(2.5), Inches(3.0),
              tbl_w - Inches(8.0)]
    row_h = Inches(0.5)

    # Header row
    add_rect(slide, M_LEFT, tbl_y, tbl_w, row_h,
             fill=AWS_DARK_NAVY, line=None)
    headers = ["시점", "점유율", "매장 수", "글로벌 벤치마크"]
    for j, h in enumerate(headers):
        _, tf = add_text(slide, cols_x[j], tbl_y, cols_w[j], row_h,
                          anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, h, size=12, bold=True, color=WHITE,
                 line_spacing=1.0, first=True)

    rows = [
        ("18m (Series A)", "3%", "약 1,700매장", "Pre-Series A SaaS 침투율"),
        ("5년 (안착)",      "10%", "약 5,700매장", "Shopify · Toast 수준"),
        ("10년 (리더)",     "20%", "약 11,400매장", "Square · Mindbody 수준"),
    ]
    for r_idx, row in enumerate(rows):
        y = tbl_y + row_h + row_h * r_idx
        fill = GRAY_050 if r_idx % 2 == 1 else WHITE
        add_rect(slide, M_LEFT, y, tbl_w, row_h,
                 fill=fill, line=GRAY_200, line_w=0.5)
        for j, cell in enumerate(row):
            bold = (j == 1 and r_idx == 2) or j == 0
            color = AWS_ORANGE_DARK if (j == 1 and r_idx == 2) else GRAY_800
            _, tf = add_text(slide, cols_x[j], y, cols_w[j], row_h,
                              anchor=MSO_ANCHOR.MIDDLE)
            add_para(tf, cell, size=11, bold=bold, color=color,
                     line_spacing=1.0, first=True)

    # Bottom caption
    _, tf = add_text(slide, M_LEFT, Inches(6.7), tbl_w, Inches(0.3))
    add_para(tf,
             "* Square이 미국 SMB POS에서 했던 일을, 또와는 한국 정육 카테고리에서 합니다.",
             size=10, italic=True, color=GRAY_500, line_spacing=1.2, first=True)


# ---------- Slide 8: Competition ----------
def s08_competition(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 8, TOTAL,
               "수직 × Zero-Task — 비어 있는 자리.",
               eyebrow="COMPETITION")

    # Left: 2x2 map
    map_x = M_LEFT
    map_y = Inches(2.3)
    map_w = Inches(5.8)
    map_h = Inches(4.3)
    add_rect(slide, map_x, map_y, map_w, map_h,
             fill=GRAY_050, line=GRAY_300, line_w=0.75, corner=True)

    # Axes
    # Y axis label (top)
    _, tf = add_text(slide, map_x + Inches(0.1), map_y + Inches(0.1),
                      Inches(2.5), Inches(0.3))
    add_para(tf, "↑ Zero-Task (사장님 손 = 0)",
             size=10, bold=True, color=AWS_ORANGE_DARK,
             line_spacing=1.0, first=True)
    # X axis label (bottom)
    _, tf = add_text(slide, map_x + Inches(0.1), map_y + map_h - Inches(0.35),
                      map_w - Inches(0.2), Inches(0.25))
    add_para(tf, "수평 (전 자영업)             ⟶             수직 (정육 특화)",
             size=10, bold=True, color=GRAY_700,
             align=PP_ALIGN.CENTER, line_spacing=1.0, first=True)

    # Quadrant dividing lines
    cx = map_x + map_w / 2
    cy = map_y + map_h / 2
    add_line(slide, cx, map_y + Inches(0.5), cx, map_y + map_h - Inches(0.5),
             color=GRAY_400, weight=0.5)
    add_line(slide, map_x + Inches(0.3), cy,
             map_x + map_w - Inches(0.3), cy,
             color=GRAY_400, weight=0.5)

    # Star: 또와 (top-right)
    star_x = map_x + Inches(3.7)
    star_y = map_y + Inches(0.9)
    add_rect(slide, star_x, star_y, Inches(1.6), Inches(0.9),
             fill=AWS_ORANGE, line=AWS_ORANGE_DARK, line_w=1.5, corner=True)
    _, tf = add_text(slide, star_x, star_y, Inches(1.6), Inches(0.9),
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "★ 또와", size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, line_spacing=1.1, first=True)
    add_para(tf, "정육 × Zero-Task",
             size=9, color=WHITE,
             align=PP_ALIGN.CENTER, line_spacing=1.0)

    # Other players (dots)
    others = [
        ("토스플레이스 CRM", map_x + Inches(0.7), map_y + Inches(1.8)),
        ("캐시노트 CRM",     map_x + Inches(0.7), map_y + Inches(2.4)),
        ("정육점 SYSTEM 대리점", map_x + Inches(3.5), map_y + Inches(3.0)),
    ]
    for name, lx, ly in others:
        # dot
        add_rect(slide, lx, ly, Inches(0.15), Inches(0.15),
                 fill=GRAY_600, line=None, corner=True)
        # label
        _, tf = add_text(slide, lx + Inches(0.22), ly - Inches(0.04),
                          Inches(2.2), Inches(0.3))
        add_para(tf, name, size=10, color=GRAY_800,
                 line_spacing=1.0, first=True)

    # Right: comparison table
    tbl_x = M_LEFT + map_w + Inches(0.3)
    tbl_y = Inches(2.3)
    tbl_w = CONTENT_W - map_w - Inches(0.3)
    cols_x = [tbl_x,
              tbl_x + Inches(2.0),
              tbl_x + Inches(4.2),
              tbl_x + Inches(5.1)]
    cols_w = [Inches(2.0), Inches(2.2), Inches(0.9),
              tbl_w - Inches(5.1)]
    row_h = Inches(0.7)

    # Header
    add_rect(slide, tbl_x, tbl_y, tbl_w, Inches(0.5),
             fill=AWS_DARK_NAVY, line=None)
    headers = ["플레이어", "산업 내 포지션", "정육 특화", "자동화"]
    for j, h in enumerate(headers):
        _, tf = add_text(slide, cols_x[j], tbl_y, cols_w[j], Inches(0.5),
                          anchor=MSO_ANCHOR.MIDDLE,
                          align=PP_ALIGN.CENTER if j > 1 else PP_ALIGN.LEFT)
        add_para(tf, h, size=10, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER if j > 1 else PP_ALIGN.LEFT,
                 line_spacing=1.0, first=True)

    rows = [
        ("또와", "시작 단계", "[O]", "[O]", True),
        ("토스플레이스 CRM", "자영업 전반 — 고객관리 우위", "[X]", "[△]", False),
        ("캐시노트 CRM", "자영업 전반 — 매출관리 우위", "[X]", "[X]", False),
        ("정육점 SYSTEM 대리점", "정육 내 점유율 우위", "[△]", "[X]", False),
    ]
    for r_idx, (player, pos, meat, auto, is_us) in enumerate(rows):
        y = tbl_y + Inches(0.5) + row_h * r_idx
        fill = AWS_ORANGE_FAINT if is_us else (GRAY_050 if r_idx % 2 == 1 else WHITE)
        line_color = AWS_ORANGE_LIGHT if is_us else GRAY_200
        add_rect(slide, tbl_x, y, tbl_w, row_h,
                 fill=fill, line=line_color, line_w=0.5)
        # Player
        _, tf = add_text(slide, cols_x[0] + Inches(0.1), y,
                          cols_w[0] - Inches(0.1), row_h,
                          anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, player,
                 size=10, bold=is_us,
                 color=AWS_ORANGE_DARK if is_us else GRAY_800,
                 line_spacing=1.2, first=True)
        # Position
        _, tf = add_text(slide, cols_x[1] + Inches(0.1), y,
                          cols_w[1] - Inches(0.1), row_h,
                          anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, pos, size=9, color=GRAY_700,
                 line_spacing=1.3, first=True)
        # Meat
        _, tf = add_text(slide, cols_x[2], y, cols_w[2], row_h,
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        meat_color = GREEN_OK if meat == "[O]" else (AMBER if meat == "[△]" else RED)
        add_para(tf, meat, size=11, bold=True, color=meat_color,
                 align=PP_ALIGN.CENTER, line_spacing=1.0, first=True)
        # Auto
        _, tf = add_text(slide, cols_x[3], y, cols_w[3], row_h,
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        auto_color = GREEN_OK if auto == "[O]" else (AMBER if auto == "[△]" else RED)
        add_para(tf, auto, size=11, bold=True, color=auto_color,
                 align=PP_ALIGN.CENTER, line_spacing=1.0, first=True)

    # Bottom caption (full width)
    _, tf = add_text(slide, M_LEFT, Inches(6.8), CONTENT_W, Inches(0.3))
    add_para(tf,
             "* 같은 정육 매장에서 또와 + 토스플레이스/캐시노트/정육점 SYSTEM 공존 가능. 결제·장부·POS는 그들이, 단골·재고·발주 자동화는 또와가.",
             size=9, italic=True, color=GRAY_500, line_spacing=1.2, first=True)


# ---------- Slide 9: Market Insight ----------
def s09_insight(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 9, TOTAL,
               "큰 경쟁자도 못 뚫은 정육점 시장.",
               eyebrow="MARKET INSIGHT",
               subtitle="이 시장이 비어있는 이유 = 우리의 해자.")

    # 3 horizontal sections: 01 사실 / 02 이유 / 03 그래서
    sec_y = Inches(2.5)
    sec_h = Inches(3.0)
    gap = Inches(0.2)
    sec_w = (CONTENT_W - gap * 2) / 3

    sections = [
        ("01", "사실",
         "토스플레이스·캐시노트도 정육점은 못 뚫었다",
         "자영업 전반엔 빠르게 침투했지만, 정육점 사장님은 안 씀.\n베타 매장 사장님 전원이 “바꿔본 적 없다” (현장 인터뷰)."),
        ("02", "이유",
         "잘 되면 안 바꾼다",
         "정육점 사장님의 보수적 충성도. 한 번 안착하면 몇 년 그대로.\n진입 장벽이 높음 = 들어간 회사에겐 시간을 벌어주는 해자."),
        ("03", "그래서",
         "가치 제안이 달라야 한다",
         "(아래 표 참조)"),
    ]
    for i, (num, lbl, head, body) in enumerate(sections):
        x = M_LEFT + (sec_w + gap) * i
        add_rect(slide, x, sec_y, sec_w, sec_h,
                 fill=WHITE, line=GRAY_300, line_w=1.0, corner=True)
        # Number badge
        add_rect(slide, x + Inches(0.3), sec_y + Inches(0.3),
                 Inches(0.5), Inches(0.5),
                 fill=AWS_ORANGE_DARK, line=None, corner=True)
        _, tf = add_text(slide, x + Inches(0.3), sec_y + Inches(0.3),
                          Inches(0.5), Inches(0.5),
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, num, size=14, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, line_spacing=1.0, first=True)
        # Label
        _, tf = add_text(slide, x + Inches(0.95), sec_y + Inches(0.35),
                          sec_w - Inches(1.2), Inches(0.4))
        add_para(tf, lbl, size=12, bold=True, color=AWS_ORANGE_DARK,
                 line_spacing=1.0, first=True)
        # Headline
        _, tf = add_text(slide, x + Inches(0.3), sec_y + Inches(0.95),
                          sec_w - Inches(0.6), Inches(0.7))
        add_para(tf, head, size=14, bold=True, color=GRAY_900,
                 line_spacing=1.3, first=True)
        # Body
        _, tf = add_text(slide, x + Inches(0.3), sec_y + Inches(1.7),
                          sec_w - Inches(0.6), sec_h - Inches(1.9))
        add_para(tf, body, size=11, color=GRAY_700, line_spacing=1.4, first=True)

    # Bottom: 03 detail comparison table
    cmp_y = sec_y + sec_h + Inches(0.25)
    cmp_h = Inches(1.2)
    cols_x = [M_LEFT, M_LEFT + Inches(2.0),
              M_LEFT + Inches(7.0)]
    cols_w = [Inches(2.0), Inches(5.0),
              CONTENT_W - Inches(7.0)]
    # Header
    add_rect(slide, M_LEFT, cmp_y, CONTENT_W, Inches(0.4),
             fill=AWS_DARK_NAVY, line=None)
    hdrs = ["", "전략", "대표 카피"]
    for j, h in enumerate(hdrs):
        _, tf = add_text(slide, cols_x[j] + Inches(0.1), cmp_y,
                          cols_w[j] - Inches(0.1), Inches(0.4),
                          anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, h, size=11, bold=True, color=WHITE,
                 line_spacing=1.0, first=True)
    # Row 1: 경쟁사
    y = cmp_y + Inches(0.4)
    add_rect(slide, M_LEFT, y, CONTENT_W, Inches(0.4),
             fill=GRAY_050, line=GRAY_200, line_w=0.5)
    _, tf = add_text(slide, cols_x[0] + Inches(0.1), y, cols_w[0], Inches(0.4),
                     anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "경쟁사 접근", size=11, bold=True, color=GRAY_800,
             line_spacing=1.0, first=True)
    _, tf = add_text(slide, cols_x[1] + Inches(0.1), y, cols_w[1], Inches(0.4),
                     anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "POS·매출관리 등 모든 오프라인 매장 대상 다기능 중심",
             size=10, color=GRAY_700, line_spacing=1.2, first=True)
    _, tf = add_text(slide, cols_x[2] + Inches(0.1), y, cols_w[2], Inches(0.4),
                     anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "“만 원 써서 2만 원 버세요” (매출 증가)",
             size=10, italic=True, color=GRAY_700, line_spacing=1.2, first=True)
    # Row 2: 또와
    y = cmp_y + Inches(0.8)
    add_rect(slide, M_LEFT, y, CONTENT_W, Inches(0.4),
             fill=AWS_ORANGE_FAINT, line=AWS_ORANGE_LIGHT, line_w=0.5)
    _, tf = add_text(slide, cols_x[0] + Inches(0.1), y, cols_w[0], Inches(0.4),
                     anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "또와 접근", size=11, bold=True, color=AWS_ORANGE_DARK,
             line_spacing=1.0, first=True)
    _, tf = add_text(slide, cols_x[1] + Inches(0.1), y, cols_w[1], Inches(0.4),
                     anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "정육 도메인 비효율 업무를 삭제, 산업 특화 편리함으로 침투",
             size=10, bold=True, color=GRAY_900, line_spacing=1.2, first=True)
    _, tf = add_text(slide, cols_x[2] + Inches(0.1), y, cols_w[2], Inches(0.4),
                     anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "“만 원 쓰던 거, 오천 원만 써도 그대로” (비용 절감)",
             size=10, italic=True, bold=True, color=AWS_ORANGE_DARK,
             line_spacing=1.2, first=True)


# ---------- Slide 10: Business Model ----------
def s10_bm(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 10, TOTAL,
               "또와는 단골관리 + 유통 중개로 양쪽에서 수익을 받는다.",
               eyebrow="BUSINESS MODEL")

    # Top: flow diagram (simplified as 5 boxes + text annotations)
    flow_y = Inches(2.3)
    flow_h = Inches(1.6)

    # 5 actors
    box_w = Inches(1.9)
    actors_x = [M_LEFT, M_LEFT + Inches(2.5), M_LEFT + Inches(5.0),
                M_LEFT + Inches(7.5), M_LEFT + Inches(10.1)]
    actors = [
        ("정육점", "(POS 포함)"),
        ("또와", "㈜위브원"),
        ("㈜더담우", "유통사"),
        ("단골 소비자", ""),
        ("정육 도매상", "(Future)"),
    ]
    for i, (name, sub) in enumerate(actors):
        x = actors_x[i]
        is_us = (i == 1)
        add_rect(slide, x, flow_y, box_w, flow_h,
                 fill=AWS_ORANGE if is_us else WHITE,
                 line=AWS_ORANGE_DARK if is_us else GRAY_400,
                 line_w=1.5 if is_us else 0.75, corner=True)
        _, tf = add_text(slide, x, flow_y, box_w, flow_h,
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, name, size=13, bold=True,
                 color=WHITE if is_us else GRAY_900,
                 align=PP_ALIGN.CENTER, line_spacing=1.2, first=True)
        if sub:
            add_para(tf, sub, size=9,
                     color=WHITE if is_us else GRAY_600,
                     align=PP_ALIGN.CENTER, line_spacing=1.2)

    # Flow legend
    leg_y = Inches(4.1)
    _, tf = add_text(slide, M_LEFT, leg_y, CONTENT_W, Inches(0.35))
    add_para(tf,
             "● 데이터 (POS→또와)   ● 가치 (또와→정육점 자동화 / 도매→정육점 배송)   ● 돈 (정육점→또와 구독+문자 / 더담우↔또와 영업 수수료)",
             size=10, color=GRAY_700, line_spacing=1.2, first=True)

    # Bottom: 3 revenue cards
    card_y = Inches(4.6)
    card_h = Inches(2.0)
    gap = Inches(0.25)
    card_w = (CONTENT_W - gap * 2) / 3

    cards = [
        ("Basic 구독", "₩33,000 / 월",
         "ARPU ₩8~12만/월\nGross Margin 75~80%"),
        ("문자 발송", "₩33 / 건",
         "매출 효과 +15% (가설)\nLTV/CAC 5배+"),
        ("유통 매입·중개", "(Future)",
         "마케팅·재고비 −50% (가설)\n월 이탈율 3%"),
    ]
    for i, (title, val, body) in enumerate(cards):
        x = M_LEFT + (card_w + gap) * i
        add_rect(slide, x, card_y, card_w, card_h,
                 fill=AWS_ORANGE_FAINT, line=AWS_ORANGE_LIGHT, line_w=1.0, corner=True)
        _, tf = add_text(slide, x + Inches(0.2), card_y + Inches(0.2),
                          card_w - Inches(0.4), Inches(0.35))
        add_para(tf, title, size=12, bold=True, color=GRAY_800,
                 line_spacing=1.0, first=True)
        _, tf = add_text(slide, x + Inches(0.2), card_y + Inches(0.6),
                          card_w - Inches(0.4), Inches(0.55))
        add_para(tf, val, size=20, bold=True, color=AWS_ORANGE_DARK,
                 line_spacing=1.0, first=True)
        _, tf = add_text(slide, x + Inches(0.2), card_y + Inches(1.2),
                          card_w - Inches(0.4), Inches(0.75))
        add_para(tf, body, size=10, color=GRAY_700, line_spacing=1.4, first=True)


# ---------- Slide 11: Traction ----------
def s11_traction(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 11, TOTAL,
               "파일럿 1매장에서 Zero-Task 워크플로 가동 중.",
               eyebrow="TRACTION")

    # Top: 3 big numbers
    num_y = Inches(2.3)
    num_h = Inches(1.6)
    gap = Inches(0.3)
    num_w = (CONTENT_W - gap * 2) / 3
    nums = [
        ("약 5만 건", "누적 거래"),
        ("약 2,100명", "누적 단골 (5회+ 방문)"),
        ("1매장", "파일럿 도입 운영 중"),
    ]
    for i, (big, lbl) in enumerate(nums):
        x = M_LEFT + (num_w + gap) * i
        add_rect(slide, x, num_y, num_w, num_h,
                 fill=AWS_ORANGE_FAINT, line=AWS_ORANGE_LIGHT, line_w=1.0, corner=True)
        _, tf = add_text(slide, x, num_y + Inches(0.25), num_w, Inches(0.7),
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, big, size=30, bold=True, color=AWS_ORANGE_DARK,
                 align=PP_ALIGN.CENTER, line_spacing=1.0, first=True)
        _, tf = add_text(slide, x, num_y + Inches(1.0), num_w, Inches(0.5),
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, lbl, size=12, color=GRAY_700,
                 align=PP_ALIGN.CENTER, line_spacing=1.2, first=True)

    # Middle: honest disclosure
    hon_y = Inches(4.2)
    hon_h = Inches(1.9)
    add_rect(slide, M_LEFT, hon_y, CONTENT_W, hon_h,
             fill=WHITE, line=AMBER, line_w=1.5, corner=True)
    _, tf = add_text(slide, M_LEFT + Inches(0.3), hon_y + Inches(0.2),
                      CONTENT_W - Inches(0.6), Inches(0.4))
    add_para(tf, "[!] 아직 측정되지 않은 것 — 정직 공개",
             size=13, bold=True, color=AMBER, line_spacing=1.0, first=True)
    _, tf = add_text(slide, M_LEFT + Inches(0.3), hon_y + Inches(0.65),
                      CONTENT_W - Inches(0.6), hon_h - Inches(0.85))
    add_para(tf, "·  도입 전/후 매장 매출 차이", size=11, color=GRAY_800,
             line_spacing=1.4, first=True)
    add_para(tf, "·  알림 발송 → 재방문 정량 효과", size=11, color=GRAY_800,
             line_spacing=1.4)
    add_para(tf, "·  매장별 잔존율·재방문 주기 단축", size=11, color=GRAY_800,
             line_spacing=1.4)
    add_para(tf, "·  제품 가치 가설 +15% / −50% 실증", size=11, color=GRAY_800,
             line_spacing=1.4)
    add_para(tf, "→ 다음 슬라이드 『Validation To-do』에서 정면으로 다룹니다.",
             size=11, italic=True, color=GRAY_500, line_spacing=1.4, space_before=4)

    # Partner
    p_y = Inches(6.3)
    _, tf = add_text(slide, M_LEFT, p_y, CONTENT_W, Inches(0.35))
    add_para(tf,
             "협력사: ㈜더담우 (연 매출 1,000억 정육 유통사)  ·  5개 점포 영업망 협력 합의 (도입 시작 전)",
             size=11, color=GRAY_700, line_spacing=1.2, first=True)


# ---------- Slide 12: Validation ----------
def s12_validation(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 12, TOTAL,
               "Pre-Seed 5억 = 가설 검증 → PMF → 공헌이익 양수화.",
               eyebrow="VALIDATION TO-DO",
               subtitle="Series A는 그 다음.")

    # 4 stages horizontal
    s_y = Inches(2.5)
    s_h = Inches(4.0)
    gap = Inches(0.2)
    s_w = (CONTENT_W - gap * 3) / 4

    stages = [
        ("Stage 1", "가설 검증", "+0~6m",
         "Q1 도입 전/후 매장 매출\nQ2 알림 → 재방문 전환율\nQ3 이탈군 잔존율",
         "통과 → Stage 2\n실패 → 제품 PIVOT",
         AWS_ORANGE),
        ("Stage 2", "PMF 찾기", "+6~12m",
         "Q4 N=30 효과 재현\nQ5 유료 전환율\nQ6 문자 발송량 분포",
         "통과 → Stage 3\n실패 → 가격·세그먼트\n     재설계",
         AWS_ORANGE_DARK),
        ("Stage 3", "공헌이익 양수화", "+12~18m",
         "Q7 더담우 영업 효율\nQ8 직판 CAC\nQ9 POS사 수락율\nQ10 실제 ARPU\nQ11 이탈율·LTV",
         "통과 → Series A 도전\n실패 → 브릿지 라운드",
         AWS_DARK_NAVY),
        ("Stage 4", "시장 확장", "Series A 이후",
         "Q12 인접 카테고리 PoC",
         "(Pre-Seed 범위 외)",
         GRAY_500),
    ]
    for i, (lbl, name, time, qs, gate, color) in enumerate(stages):
        x = M_LEFT + (s_w + gap) * i
        # Top color bar
        add_rect(slide, x, s_y, s_w, Inches(0.1), fill=color, line=None)
        # Card
        add_rect(slide, x, s_y + Inches(0.1), s_w, s_h - Inches(0.1),
                 fill=WHITE, line=GRAY_300, line_w=1.0, corner=True)
        # Label
        _, tf = add_text(slide, x + Inches(0.2), s_y + Inches(0.25),
                          s_w - Inches(0.4), Inches(0.35))
        add_para(tf, lbl, size=11, bold=True, color=color,
                 line_spacing=1.0, first=True)
        # Name
        _, tf = add_text(slide, x + Inches(0.2), s_y + Inches(0.65),
                          s_w - Inches(0.4), Inches(0.45))
        add_para(tf, name, size=15, bold=True, color=GRAY_900,
                 line_spacing=1.1, first=True)
        # Time
        _, tf = add_text(slide, x + Inches(0.2), s_y + Inches(1.15),
                          s_w - Inches(0.4), Inches(0.3))
        add_para(tf, time, size=10, color=GRAY_500,
                 line_spacing=1.0, first=True)
        # Questions
        _, tf = add_text(slide, x + Inches(0.2), s_y + Inches(1.55),
                          s_w - Inches(0.4), Inches(1.6))
        add_para(tf, qs, size=10, color=GRAY_800, line_spacing=1.4, first=True)
        # Gate
        gate_y = s_y + s_h - Inches(0.95)
        add_rect(slide, x + Inches(0.15), gate_y,
                 s_w - Inches(0.3), Inches(0.8),
                 fill=AWS_ORANGE_FAINT, line=AWS_ORANGE_LIGHT, line_w=0.5, corner=True)
        _, tf = add_text(slide, x + Inches(0.25), gate_y,
                          s_w - Inches(0.5), Inches(0.8),
                          anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, gate, size=9, color=GRAY_800, line_spacing=1.3, first=True)

    # Bottom: decision gates
    g_y = Inches(6.7)
    _, tf = add_text(slide, M_LEFT, g_y, CONTENT_W, Inches(0.3))
    add_para(tf,
             "의사결정 게이트  —  +6m: Stage 1 → 베타 가속   ·   +12m: Stage 2 → 유료 GTM   ·   +18m: Stage 3 → Series A",
             size=10, italic=True, color=GRAY_600, line_spacing=1.2, first=True)


# ---------- Slide 13: Team ----------
def s13_team(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 13, TOTAL,
               "B2B SaaS + AI 박사 + 정육 도메인 — 한 팀.",
               eyebrow="TEAM")

    # 5 person cards (horizontal)
    p_y = Inches(2.4)
    p_h = Inches(3.5)
    gap = Inches(0.15)
    p_w = (CONTENT_W - gap * 4) / 5

    team = [
        ("안동철", "대표이사", "한밭대 컴공",
         "㈜스낵포 CTO 8년\nB2B SaaS 0→1"),
        ("배현혜", "AI·데이터 총괄", "Columbia 박사",
         "빅데이터 전임연구원\n㈜팬블러 창업"),
        ("SEO CHARLES", "제품·UX 총괄", "KAIST 학·석사",
         "㈜팬블러\nCo-Founder/PO"),
        ("이지백", "현장·도메인", "충남대 경영 박사",
         "정육 유통사\n사외이사"),
        ("박동일", "운영·재무", "한밭대 회계학사",
         "—"),
    ]
    for i, (name, role, edu, exp) in enumerate(team):
        x = M_LEFT + (p_w + gap) * i
        add_rect(slide, x, p_y, p_w, p_h,
                 fill=WHITE, line=GRAY_300, line_w=1.0, corner=True)
        # Photo placeholder circle
        photo_size = Inches(1.0)
        photo_x = x + (p_w - photo_size) / 2
        add_rect(slide, photo_x, p_y + Inches(0.3),
                 photo_size, photo_size,
                 fill=AWS_ORANGE_LIGHT, line=AWS_ORANGE_LIGHT, corner=True)
        # Name
        _, tf = add_text(slide, x, p_y + Inches(1.4), p_w, Inches(0.4),
                          align=PP_ALIGN.CENTER)
        add_para(tf, name, size=14, bold=True, color=GRAY_900,
                 align=PP_ALIGN.CENTER, line_spacing=1.1, first=True)
        # Role
        _, tf = add_text(slide, x, p_y + Inches(1.8), p_w, Inches(0.35),
                          align=PP_ALIGN.CENTER)
        add_para(tf, role, size=10, bold=True, color=AWS_ORANGE_DARK,
                 align=PP_ALIGN.CENTER, line_spacing=1.1, first=True)
        # Edu
        _, tf = add_text(slide, x + Inches(0.1), p_y + Inches(2.2),
                          p_w - Inches(0.2), Inches(0.4),
                          align=PP_ALIGN.CENTER)
        add_para(tf, edu, size=9, italic=True, color=GRAY_500,
                 align=PP_ALIGN.CENTER, line_spacing=1.2, first=True)
        # Experience
        _, tf = add_text(slide, x + Inches(0.1), p_y + Inches(2.6),
                          p_w - Inches(0.2), Inches(0.85),
                          align=PP_ALIGN.CENTER)
        add_para(tf, exp, size=9, color=GRAY_700,
                 align=PP_ALIGN.CENTER, line_spacing=1.3, first=True)

    # Bottom Founder-Market Fit strip
    fmf_y = Inches(6.1)
    add_rect(slide, M_LEFT, fmf_y, CONTENT_W, Inches(0.7),
             fill=AWS_ORANGE_FAINT, line=AWS_ORANGE_LIGHT, line_w=1.0, corner=True)
    _, tf = add_text(slide, M_LEFT, fmf_y, CONTENT_W, Inches(0.7),
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "Founder-Market Fit",
             size=11, bold=True, color=AWS_ORANGE_DARK,
             align=PP_ALIGN.CENTER, line_spacing=1.1, first=True)
    add_para(tf, "B2B SaaS 실행력 + AI 박사급 + 정육 도메인 네트워크 — 한 팀 안에 3축 모두",
             size=10, color=GRAY_700,
             align=PP_ALIGN.CENTER, line_spacing=1.2)


# ---------- Slide 14: Roadmap ----------
def s14_roadmap(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 14, TOTAL,
               "Pre-Seed 5억 + 정부 R&D 7억 / 18개월 — 공헌이익 양수까지 → Series A.",
               eyebrow="ROADMAP")

    # Timeline horizontal
    tl_y = Inches(2.6)
    tl_h = Inches(2.6)

    # Stage markers (5 milestones)
    mks = [
        ("NOW", "파일럿 1", ""),
        ("+6m", "베타 N=30", "Stage 1"),
        ("+12m", "유료 100~150", "Stage 2"),
        ("+18m", "Series A 도전", "Stage 3"),
        ("이후", "시장 확장", "Stage 4"),
    ]
    n = len(mks)
    avail_w = CONTENT_W
    step = avail_w / (n - 1)
    line_y = tl_y + Inches(0.5)

    # Horizontal line
    add_line(slide, M_LEFT, line_y,
             M_LEFT + avail_w, line_y,
             color=AWS_ORANGE, weight=2.5)

    for i, (time, milestone, stage) in enumerate(mks):
        x = M_LEFT + step * i
        # Dot
        dot_size = Inches(0.25)
        add_rect(slide, x - dot_size / 2, line_y - dot_size / 2,
                 dot_size, dot_size,
                 fill=AWS_ORANGE_DARK, line=WHITE, line_w=2.0, corner=True)
        # Time above
        _, tf = add_text(slide, x - Inches(1.0), line_y - Inches(0.6),
                          Inches(2.0), Inches(0.3),
                          align=PP_ALIGN.CENTER)
        add_para(tf, time, size=11, bold=True, color=AWS_ORANGE_DARK,
                 align=PP_ALIGN.CENTER, line_spacing=1.0, first=True)
        # Milestone below
        _, tf = add_text(slide, x - Inches(1.2), line_y + Inches(0.2),
                          Inches(2.4), Inches(0.4),
                          align=PP_ALIGN.CENTER)
        add_para(tf, milestone, size=12, bold=True, color=GRAY_900,
                 align=PP_ALIGN.CENTER, line_spacing=1.2, first=True)
        # Stage
        if stage:
            _, tf = add_text(slide, x - Inches(1.0), line_y + Inches(0.65),
                              Inches(2.0), Inches(0.3),
                              align=PP_ALIGN.CENTER)
            add_para(tf, stage, size=10, italic=True, color=GRAY_500,
                     align=PP_ALIGN.CENTER, line_spacing=1.0, first=True)

    # Funding bands
    band_y = Inches(4.5)
    bw1 = Inches(8.0)  # Pre-Seed
    add_rect(slide, M_LEFT, band_y, bw1, Inches(0.45),
             fill=AWS_ORANGE_LIGHT, line=AWS_ORANGE, line_w=1.0, corner=True)
    _, tf = add_text(slide, M_LEFT + Inches(0.2), band_y,
                      bw1 - Inches(0.4), Inches(0.45),
                      anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "Pre-Seed 5억 — 희석 자금, GTM·영업·운영",
             size=11, bold=True, color=AWS_ORANGE_DARK,
             line_spacing=1.0, first=True)

    band_y2 = band_y + Inches(0.55)
    bw2 = Inches(8.0)
    add_rect(slide, M_LEFT, band_y2, bw2, Inches(0.45),
             fill=AWS_ORANGE_FAINT, line=AWS_ORANGE_LIGHT, line_w=1.0, corner=True)
    _, tf = add_text(slide, M_LEFT + Inches(0.2), band_y2,
                      bw2 - Inches(0.4), Inches(0.45),
                      anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "정부 R&D 7억 — 비희석, 디딤돌 + TIPS + 후속 (가설)",
             size=11, bold=True, color=GRAY_800,
             line_spacing=1.0, first=True)

    # Series A entry box
    sa_x = M_LEFT + bw1 + Inches(0.3)
    sa_w = CONTENT_W - bw1 - Inches(0.3)
    add_rect(slide, sa_x, band_y, sa_w, Inches(1.0),
             fill=AWS_DARK_NAVY, line=AWS_DARK_NAVY, corner=True)
    _, tf = add_text(slide, sa_x, band_y, sa_w, Inches(1.0),
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "Series A →", size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, line_spacing=1.1, first=True)
    add_para(tf, "공헌이익 양수\n유료 매장 100~150\nLTV/CAC 3~5배",
             size=9, color=WHITE,
             align=PP_ALIGN.CENTER, line_spacing=1.3)

    # Bottom Series A entry conditions
    sec_y = Inches(6.0)
    _, tf = add_text(slide, M_LEFT, sec_y, CONTENT_W, Inches(0.3))
    add_para(tf, "Series A 진입 조건 (+18m 목표)",
             size=11, bold=True, color=AWS_ORANGE_DARK,
             line_spacing=1.0, first=True)
    cond_y = sec_y + Inches(0.35)
    conds = ["공헌이익 양수", "유료 매장 100~150", "LTV/CAC 3~5배",
             "월 이탈율 <5%", "협력 채널 더담우 + POS사"]
    cw = CONTENT_W / len(conds)
    for i, c in enumerate(conds):
        x = M_LEFT + cw * i
        add_rect(slide, x + Inches(0.05), cond_y,
                 cw - Inches(0.1), Inches(0.4),
                 fill=WHITE, line=GRAY_300, line_w=0.75, corner=True)
        _, tf = add_text(slide, x + Inches(0.05), cond_y,
                          cw - Inches(0.1), Inches(0.4),
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, c, size=10, color=GRAY_800,
                 align=PP_ALIGN.CENTER, line_spacing=1.1, first=True)


# ---------- Slide 15: The Ask ----------
def s15_ask(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 15, TOTAL,
               "Pre-Seed 5억 + R&D 7억 — 18개월, 공헌이익 양수까지.",
               eyebrow="THE ASK")

    # Top: 2 big numbers
    h_y = Inches(2.3)
    h_h = Inches(1.7)
    gap = Inches(0.3)
    h_w = (CONTENT_W - gap) / 2

    # Pre-Seed
    add_rect(slide, M_LEFT, h_y, h_w, h_h,
             fill=AWS_ORANGE, line=AWS_ORANGE_DARK, line_w=1.5, corner=True)
    _, tf = add_text(slide, M_LEFT + Inches(0.4), h_y + Inches(0.3),
                      h_w - Inches(0.8), Inches(0.5))
    add_para(tf, "Pre-Seed", size=14, bold=True, color=WHITE,
             line_spacing=1.0, first=True)
    _, tf = add_text(slide, M_LEFT + Inches(0.4), h_y + Inches(0.8),
                      h_w - Inches(0.8), Inches(0.65))
    add_para(tf, "5억", size=44, bold=True, color=WHITE,
             line_spacing=1.0, first=True)
    _, tf = add_text(slide, M_LEFT + Inches(0.4), h_y + h_h - Inches(0.45),
                      h_w - Inches(0.8), Inches(0.35))
    add_para(tf, "희석 자금", size=11, italic=True, color=WHITE,
             line_spacing=1.0, first=True)

    # R&D
    x2 = M_LEFT + h_w + gap
    add_rect(slide, x2, h_y, h_w, h_h,
             fill=AWS_DARK_NAVY, line=AWS_DARK_NAVY, corner=True)
    _, tf = add_text(slide, x2 + Inches(0.4), h_y + Inches(0.3),
                      h_w - Inches(0.8), Inches(0.5))
    add_para(tf, "정부 R&D", size=14, bold=True, color=WHITE,
             line_spacing=1.0, first=True)
    _, tf = add_text(slide, x2 + Inches(0.4), h_y + Inches(0.8),
                      h_w - Inches(0.8), Inches(0.65))
    add_para(tf, "7억", size=44, bold=True, color=WHITE,
             line_spacing=1.0, first=True)
    _, tf = add_text(slide, x2 + Inches(0.4), h_y + h_h - Inches(0.45),
                      h_w - Inches(0.8), Inches(0.35))
    add_para(tf, "비희석 — 디딤돌 2.668억 확보 + TIPS·후속 가설",
             size=11, italic=True, color=WHITE,
             line_spacing=1.0, first=True)

    # Bottom: usage table
    u_y = Inches(4.3)
    u_h = Inches(2.2)
    u_w = CONTENT_W
    _, tf = add_text(slide, M_LEFT, Inches(4.05), u_w, Inches(0.3))
    add_para(tf, "자금 사용 계획 — 총 12억 통합",
             size=12, bold=True, color=AWS_ORANGE_DARK,
             line_spacing=1.0, first=True)

    cols_x = [M_LEFT, M_LEFT + Inches(2.5),
              M_LEFT + Inches(4.0), M_LEFT + Inches(5.5)]
    cols_w = [Inches(2.5), Inches(1.5), Inches(1.5),
              u_w - Inches(5.5)]
    row_h = Inches(0.42)

    # Header
    add_rect(slide, M_LEFT, u_y, u_w, row_h,
             fill=AWS_DARK_NAVY, line=None)
    hdrs = ["항목", "금액", "비율", "용도"]
    for j, h in enumerate(hdrs):
        _, tf = add_text(slide, cols_x[j] + Inches(0.1), u_y,
                          cols_w[j] - Inches(0.1), row_h,
                          anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, h, size=11, bold=True, color=WHITE,
                 line_spacing=1.0, first=True)

    rows = [
        ("인력", "5.0억", "42%", "영업·운영 + R&D·개발"),
        ("마케팅·영업", "3.0억", "25%", "GTM · 정육 협회 · 바우처 채널"),
        ("제품 개발", "3.0억", "25%", "AI 모델 · POS 연동 · SaaS화 (R&D 자금)"),
        ("기타 운영비", "1.0억", "8%", "사무실 · 법무 · 회계"),
        ("합계", "12.0억", "100%", "18개월 Runway"),
    ]
    for r_idx, row in enumerate(rows):
        is_total = (r_idx == len(rows) - 1)
        y = u_y + row_h + (row_h - Inches(0.05)) * r_idx
        fill = AWS_ORANGE_FAINT if is_total else (GRAY_050 if r_idx % 2 == 1 else WHITE)
        add_rect(slide, M_LEFT, y, u_w, row_h - Inches(0.05),
                 fill=fill, line=GRAY_200, line_w=0.5)
        for j, cell in enumerate(row):
            color = AWS_ORANGE_DARK if is_total else GRAY_800
            _, tf = add_text(slide, cols_x[j] + Inches(0.1), y,
                              cols_w[j] - Inches(0.1), row_h - Inches(0.05),
                              anchor=MSO_ANCHOR.MIDDLE)
            add_para(tf, cell, size=10, bold=is_total or j == 0,
                     color=color, line_spacing=1.1, first=True)

    # Bottom: decision gates
    g_y = Inches(6.8)
    _, tf = add_text(slide, M_LEFT, g_y, CONTENT_W, Inches(0.3))
    add_para(tf,
             "의사결정 게이트 — +6m: Stage 1 → 베타 가속  ·  +12m: Stage 2 → 유료 GTM  ·  +18m: Stage 3 → Series A 도전",
             size=10, italic=True, color=GRAY_600, line_spacing=1.2, first=True)


# ---------- Slide 16: Vision ----------
def s16_vision(prs):
    slide = slide_blank(prs)
    # Special layout: no chrome, more dramatic
    # Orange band on left
    add_rect(slide, 0, 0, Inches(0.3), SLIDE_H, fill=AWS_ORANGE, line=None)

    # Eyebrow
    _, tf = add_text(slide, Inches(1.0), Inches(1.0), Inches(11), Inches(0.4))
    add_para(tf, "VISION", size=12, bold=True, color=AWS_ORANGE,
             line_spacing=1.0, first=True)

    # Big quote
    _, tf = add_text(slide, Inches(1.0), Inches(1.8), Inches(11.5), Inches(4.0))
    add_para(tf, "기술 발전의 혜택에서 소외된",
             size=32, bold=True, color=GRAY_900, line_spacing=1.4, first=True)
    add_para(tf, "소상공인의 도메인에 깊숙이 파고들어",
             size=32, bold=True, color=GRAY_900, line_spacing=1.4)
    add_para(tf, "기술과의 거리감을 줄이는 비즈니스를",
             size=32, bold=True, color=GRAY_900, line_spacing=1.4)
    add_para(tf, "하겠습니다.",
             size=32, bold=True, color=AWS_ORANGE_DARK, line_spacing=1.4)

    # Closing line
    _, tf = add_text(slide, Inches(1.0), Inches(6.0), Inches(11.5), Inches(0.5))
    add_para(tf, "— 정육점은 그 시작입니다.",
             size=18, italic=True, color=GRAY_600, line_spacing=1.2, first=True)

    # Footer
    add_line(slide, Inches(1.0), Inches(6.7), Inches(6.0), Inches(6.7),
             color=AWS_ORANGE, weight=1.5)
    _, tf = add_text(slide, Inches(1.0), Inches(6.8), Inches(10), Inches(0.4))
    add_para(tf, "㈜위브원  ·  또와  ·  안동철 대표",
             size=12, color=GRAY_600, line_spacing=1.2, first=True)

    # Page number
    _, tf = add_text(slide, SLIDE_W - Inches(1.6), PAGE_NUM_TOP,
                     Inches(1.3), Inches(0.3), align=PP_ALIGN.RIGHT)
    add_para(tf, f"16 / {TOTAL}", size=10, color=GRAY_400,
             line_spacing=1.0, align=PP_ALIGN.RIGHT, first=True)


# ============================================================
# Main
# ============================================================

def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    builders = [
        s01_cover, s02_oneliner, s03_problem, s04_solution,
        s05_concept, s06_whynow, s07_market, s08_competition,
        s09_insight, s10_bm, s11_traction, s12_validation,
        s13_team, s14_roadmap, s15_ask, s16_vision,
    ]
    for b in builders:
        b(prs)

    out_dir = Path("/tmp/butcher-build")
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / "또와_IR덱_v1.2_AWS.pptx"
    prs.save(out)
    print(f"[OK] saved → {out}")
    print(f"     slides: {len(prs.slides)} / target {TOTAL}")


if __name__ == "__main__":
    build()
