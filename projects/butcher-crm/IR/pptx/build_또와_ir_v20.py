#!/usr/bin/env python3
"""
또와 IR 덱 v2.0 — PowerPoint builder (AWS tone, 17 slides, info-dense)

Aligns with: projects/butcher-crm/IR/또와_IR덱_v2.0.md
Output: /tmp/butcher-build/또와_IR덱_v2.0_AWS.pptx

Anti-breakage 5 principles (same as v12):
  1. Font 3-slot forced (Noto Sans KR)
  2. No emojis (text/shape replacements)
  3. word_wrap + sized boxes + line_spacing
  4. Simple boxes + straight connectors
  5. Manual boxes (no native tables)
"""

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ============================================================
# Design system — AWS tone
# ============================================================
AWS_ORANGE       = RGBColor(0xD1, 0x51, 0x27)
AWS_ORANGE_DARK  = RGBColor(0xA8, 0x40, 0x1E)
AWS_ORANGE_LIGHT = RGBColor(0xFA, 0xD9, 0xCB)
AWS_ORANGE_FAINT = RGBColor(0xFF, 0xF7, 0xF2)
AWS_ORANGE_MID   = RGBColor(0xE8, 0x94, 0x76)
AWS_DARK_NAVY    = RGBColor(0x23, 0x2F, 0x3E)

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

GREEN_OK = RGBColor(0x16, 0xA3, 0x4A)
AMBER    = RGBColor(0xF5, 0x9E, 0x0B)
RED      = RGBColor(0xDC, 0x26, 0x26)

FONT_KR = "Noto Sans KR"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
M_LEFT  = Inches(0.6)
M_RIGHT = Inches(0.6)
CONTENT_W = SLIDE_W - M_LEFT - M_RIGHT
TITLE_TOP = Inches(0.45)
PAGE_NUM_TOP = Inches(7.08)
TOTAL = 17


# ============================================================
# Helpers
# ============================================================
def _force_font_xml(run):
    rPr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        for el in rPr.findall(qn(f"a:{tag}")):
            rPr.remove(el)
        el = etree.SubElement(rPr, qn(f"a:{tag}"))
        el.set("typeface", FONT_KR)


def set_run(run, text, *, size=14, bold=False, italic=False, color=GRAY_800):
    run.text = text
    run.font.name = FONT_KR
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    _force_font_xml(run)


def add_rect(slide, left, top, width, height, *, fill=WHITE, line=None,
             line_w=0.75, corner=False):
    st = MSO_SHAPE.ROUNDED_RECTANGLE if corner else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(st, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_w)
    try:
        shp.shadow.inherit = False
    except Exception:
        pass
    if corner:
        try:
            shp.adjustments[0] = 0.06
        except Exception:
            pass
    return shp


def add_text(slide, left, top, width, height, *,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, word_wrap=True):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = word_wrap
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.03)
    tf.vertical_anchor = anchor
    tf.paragraphs[0].alignment = align
    return tb, tf


def add_para(tf, text="", *, size=14, bold=False, italic=False, color=GRAY_800,
             align=PP_ALIGN.LEFT, line_spacing=1.25, space_before=0,
             space_after=0, first=False):
    if first:
        p = tf.paragraphs[0]
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
    conn = slide.shapes.add_connector(1, x1, y1, x2 - x1, y2 - y1)
    conn.line.color.rgb = color
    conn.line.width = Pt(weight)
    return conn


def slide_blank(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill=WHITE, line=None)
    return slide


def add_chrome(slide, page_num, title_text, *, eyebrow=None, subtitle=None):
    add_rect(slide, M_LEFT, TITLE_TOP, Inches(0.08), Inches(0.38),
             fill=AWS_ORANGE, line=None)
    title_y = TITLE_TOP
    if eyebrow:
        _, tf = add_text(slide, M_LEFT + Inches(0.2), TITLE_TOP - Inches(0.04),
                         CONTENT_W, Inches(0.28))
        add_para(tf, eyebrow, size=10, bold=True, color=AWS_ORANGE,
                 line_spacing=1.0, first=True)
        title_y = TITLE_TOP + Inches(0.22)
    _, tf = add_text(slide, M_LEFT + Inches(0.2), title_y,
                     CONTENT_W - Inches(0.2), Inches(0.55))
    add_para(tf, title_text, size=21, bold=True, color=GRAY_900,
             line_spacing=1.12, first=True)
    if subtitle:
        _, tf = add_text(slide, M_LEFT + Inches(0.2), title_y + Inches(0.5),
                         CONTENT_W, Inches(0.35))
        add_para(tf, subtitle, size=12, italic=True, color=GRAY_600,
                 line_spacing=1.15, first=True)
    _, tf = add_text(slide, SLIDE_W - Inches(1.6), PAGE_NUM_TOP,
                     Inches(1.3), Inches(0.3), align=PP_ALIGN.RIGHT)
    add_para(tf, f"{page_num} / {TOTAL}", size=10, color=GRAY_400,
             line_spacing=1.0, align=PP_ALIGN.RIGHT, first=True)


def page_num_only(slide, n):
    _, tf = add_text(slide, SLIDE_W - Inches(1.6), PAGE_NUM_TOP,
                     Inches(1.3), Inches(0.3), align=PP_ALIGN.RIGHT)
    add_para(tf, f"{n} / {TOTAL}", size=10, color=GRAY_400,
             line_spacing=1.0, align=PP_ALIGN.RIGHT, first=True)


# Generic manual-table renderer (info-dense friendly)
def table(slide, x, y, total_w, col_w, rows, *, header=None,
          header_fill=AWS_DARK_NAVY, row_h=Inches(0.45),
          hl_row=None, hl_fill=AWS_ORANGE_FAINT, hl_line=AWS_ORANGE_LIGHT,
          font=10, header_font=10, aligns=None):
    """rows: list of tuples; col_w: list of widths summing ~total_w."""
    n_col = len(col_w)
    if aligns is None:
        aligns = [PP_ALIGN.LEFT] * n_col
    xs = [x]
    for w in col_w[:-1]:
        xs.append(xs[-1] + w)
    cur_y = y
    if header:
        add_rect(slide, x, cur_y, total_w, row_h, fill=header_fill, line=None)
        for j, h in enumerate(header):
            _, tf = add_text(slide, xs[j] + Inches(0.08), cur_y,
                             col_w[j] - Inches(0.1), row_h,
                             anchor=MSO_ANCHOR.MIDDLE, align=aligns[j])
            add_para(tf, h, size=header_font, bold=True, color=WHITE,
                     align=aligns[j], line_spacing=1.05, first=True)
        cur_y += row_h
    for r_idx, row in enumerate(rows):
        is_hl = (hl_row is not None and r_idx == hl_row)
        fill = hl_fill if is_hl else (GRAY_050 if r_idx % 2 == 1 else WHITE)
        line = hl_line if is_hl else GRAY_200
        add_rect(slide, x, cur_y, total_w, row_h, fill=fill, line=line, line_w=0.5)
        for j, cell in enumerate(row):
            color = AWS_ORANGE_DARK if is_hl else GRAY_800
            bold = is_hl or (j == 0)
            _, tf = add_text(slide, xs[j] + Inches(0.08), cur_y,
                             col_w[j] - Inches(0.1), row_h,
                             anchor=MSO_ANCHOR.MIDDLE, align=aligns[j])
            add_para(tf, cell, size=font, bold=bold, color=color,
                     align=aligns[j], line_spacing=1.15, first=True)
        cur_y += row_h
    return cur_y


# ============================================================
# Slides
# ============================================================

# ---- 1. Cover ----
def s01(prs):
    slide = slide_blank(prs)
    add_rect(slide, 0, 0, Inches(0.3), SLIDE_H, fill=AWS_ORANGE, line=None)
    _, tf = add_text(slide, Inches(1.0), Inches(1.3), Inches(10), Inches(0.4))
    add_para(tf, "㈜위브원  |  또와", size=14, bold=True, color=AWS_ORANGE,
             line_spacing=1.0, first=True)
    _, tf = add_text(slide, Inches(1.0), Inches(2.0), Inches(11.5), Inches(2.0))
    add_para(tf, "정육점 특화", size=40, color=GRAY_900, line_spacing=1.1, first=True)
    add_para(tf, "Zero-Task AI 에이전트", size=50, bold=True, color=AWS_ORANGE_DARK,
             line_spacing=1.1)
    _, tf = add_text(slide, Inches(1.0), Inches(4.4), Inches(11.5), Inches(0.6))
    add_para(tf, "토스도, 캐시노트도 못 뚫은 시장. 우리는 정육에서 시작합니다.",
             size=18, italic=True, color=GRAY_700, line_spacing=1.2, first=True)
    add_line(slide, Inches(1.0), Inches(5.9), Inches(6.0), Inches(5.9),
             color=AWS_ORANGE, weight=1.5)
    _, tf = add_text(slide, Inches(1.0), Inches(6.0), Inches(8), Inches(0.6))
    add_para(tf, "발표: 안동철  ·  대표이사  ·  2026.05",
             size=13, color=GRAY_600, line_spacing=1.2, first=True)
    page_num_only(slide, 1)


# ---- 2. Hook ----
def s02(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 2, "토스플레이스도, 캐시노트도 — 정육점은 못 뚫었습니다.",
               eyebrow="HOOK")
    # 2 fact cards
    fy = Inches(2.0)
    fh = Inches(1.5)
    fw = (CONTENT_W - Inches(0.3)) / 2
    facts = [
        ("자영업 전반엔 빠르게 침투한 거대 SaaS들이,",
         "유독 정육점에서는 점유율이 거의 0."),
        ("베타 매장 사장님 전원이",
         "“그런 거 바꿔본 적 없다” (현장 인터뷰)."),
    ]
    for i, (a, b) in enumerate(facts):
        x = M_LEFT + (fw + Inches(0.3)) * i
        add_rect(slide, x, fy, fw, fh, fill=AWS_ORANGE_FAINT,
                 line=AWS_ORANGE_LIGHT, line_w=1.0, corner=True)
        _, tf = add_text(slide, x + Inches(0.35), fy, fw - Inches(0.7), fh,
                         anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, a, size=15, color=GRAY_800, line_spacing=1.3, first=True)
        add_para(tf, b, size=15, bold=True, color=AWS_ORANGE_DARK, line_spacing=1.3)

    # Two questions block
    qy = Inches(3.9)
    add_rect(slide, M_LEFT, qy, CONTENT_W, Inches(2.0),
             fill=AWS_DARK_NAVY, line=None, corner=True)
    _, tf = add_text(slide, M_LEFT + Inches(0.5), qy + Inches(0.3),
                     CONTENT_W - Inches(1.0), Inches(1.4),
                     anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "질문은 두 가지입니다.", size=16, bold=True, color=WHITE,
             line_spacing=1.4, first=True)
    add_para(tf, "①  왜 아무도 못 뚫었나?", size=22, bold=True, color=AWS_ORANGE_LIGHT,
             line_spacing=1.5)
    add_para(tf, "②  그런데 왜 또와는 뚫을 수 있나?", size=22, bold=True,
             color=AWS_ORANGE_LIGHT, line_spacing=1.5)

    _, tf = add_text(slide, M_LEFT, qy + Inches(2.1), CONTENT_W, Inches(0.4))
    add_para(tf, "이 덱은 이 두 질문에 답합니다.", size=13, italic=True,
             color=GRAY_600, line_spacing=1.2, first=True)


# ---- 3. Why no one cracked it ----
def s03(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 3, "정육은 다르게 움직이는 시장 — 그래서 범용 SaaS가 못 들어옵니다.",
               eyebrow="WHY NO ONE CRACKED IT", subtitle="진입 장벽이 높다 = 먼저 들어간 회사에게 시간을 벌어주는 해자.")
    # Left: 3 reasons
    lx = M_LEFT
    lw = Inches(7.3)
    ry = Inches(2.4)
    rh = Inches(1.3)
    reasons = [
        ("01", "정육 시장의 비정형성",
         "판매자 협상력이 높고 거래가 비정형적(부분육·가구단위·명절·외상). 범용 POS·CRM의 표준 기능으론 못 담음."),
        ("02", "사장님의 보수적 충성도",
         "“잘 되면 안 바꾼다.” 한 번 안착하면 몇 년 그대로. 신규 도구 진입 장벽이 매우 높음."),
        ("03", "경쟁사의 구조적 한계 (못 하는 게 아니라 안 함)",
         "토스·캐시노트 ROI는 전 자영업 공통 기능에서 나옴. 식당·카페가 정육의 N배 → 정육 도메인 학습은 영원히 후순위."),
    ]
    for i, (num, t, b) in enumerate(reasons):
        y = ry + (rh + Inches(0.12)) * i
        add_rect(slide, lx, y, lw, rh, fill=WHITE, line=GRAY_300, line_w=1.0, corner=True)
        add_rect(slide, lx, y, Inches(0.08), rh, fill=AWS_ORANGE, line=None)
        _, tf = add_text(slide, lx + Inches(0.25), y + Inches(0.12),
                         Inches(0.6), Inches(0.4))
        add_para(tf, num, size=18, bold=True, color=AWS_ORANGE_DARK,
                 line_spacing=1.0, first=True)
        _, tf = add_text(slide, lx + Inches(0.9), y + Inches(0.12),
                         lw - Inches(1.1), Inches(0.4))
        add_para(tf, t, size=14, bold=True, color=GRAY_900, line_spacing=1.2, first=True)
        _, tf = add_text(slide, lx + Inches(0.9), y + Inches(0.55),
                         lw - Inches(1.1), Inches(0.7))
        add_para(tf, b, size=11, color=GRAY_700, line_spacing=1.35, first=True)

    # Right: 2x2 map
    mx = M_LEFT + lw + Inches(0.3)
    mw = CONTENT_W - lw - Inches(0.3)
    my = Inches(2.4)
    mh = Inches(4.0)
    add_rect(slide, mx, my, mw, mh, fill=GRAY_050, line=GRAY_300, line_w=0.75, corner=True)
    _, tf = add_text(slide, mx + Inches(0.15), my + Inches(0.1), mw - Inches(0.3), Inches(0.3))
    add_para(tf, "↑ Zero-Task", size=10, bold=True, color=AWS_ORANGE_DARK,
             line_spacing=1.0, first=True)
    cx = mx + mw / 2
    cy = my + mh / 2
    add_line(slide, cx, my + Inches(0.45), cx, my + mh - Inches(0.45), color=GRAY_400, weight=0.5)
    add_line(slide, mx + Inches(0.25), cy, mx + mw - Inches(0.25), cy, color=GRAY_400, weight=0.5)
    # star
    sx = mx + mw - Inches(1.7)
    sy = my + Inches(0.7)
    add_rect(slide, sx, sy, Inches(1.5), Inches(0.8), fill=AWS_ORANGE,
             line=AWS_ORANGE_DARK, line_w=1.5, corner=True)
    _, tf = add_text(slide, sx, sy, Inches(1.5), Inches(0.8),
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "★ 또와", size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, line_spacing=1.0, first=True)
    others = [("토스플레이스", mx + Inches(0.4), my + Inches(1.6)),
              ("캐시노트", mx + Inches(0.4), my + Inches(2.1)),
              ("정육점 SYSTEM\n대리점", mx + Inches(2.7), my + Inches(2.6))]
    for name, lxx, lyy in others:
        add_rect(slide, lxx, lyy, Inches(0.13), Inches(0.13), fill=GRAY_600, line=None, corner=True)
        _, tf = add_text(slide, lxx + Inches(0.2), lyy - Inches(0.05), Inches(2.0), Inches(0.45))
        add_para(tf, name, size=9, color=GRAY_800, line_spacing=1.05, first=True)
    _, tf = add_text(slide, mx + Inches(0.15), my + mh - Inches(0.35),
                     mw - Inches(0.3), Inches(0.25))
    add_para(tf, "수평 ───→ 수직(정육)", size=9, bold=True, color=GRAY_700,
             align=PP_ALIGN.CENTER, line_spacing=1.0, first=True)

    # bottom callout
    _, tf = add_text(slide, M_LEFT, Inches(6.55), CONTENT_W, Inches(0.4))
    add_para(tf, "→ 그래서 이 시장은 비어 있습니다. 우리는 정육 도메인 완전 특화로 이 견고한 시장을 뚫습니다.",
             size=12, bold=True, italic=True, color=AWS_ORANGE_DARK, line_spacing=1.2, first=True)


# ---- 4. Problem ----
def s04(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 4, "정육점 매출의 70%는 단골에서. 그런데 사장님은 단골을 잃는 줄도 모릅니다.",
               eyebrow="PROBLEM")
    cy = Inches(2.5)
    ch = Inches(3.2)
    gap = Inches(0.25)
    cw = (CONTENT_W - gap * 2) / 3
    items = [
        ("01", "단골 정보가 사장님 머릿속에만 있다",
         "파일럿 1매장만 봐도 단골 약 2,100명. 사장님이 이름·패턴을 인지하는 단골은 극소수."),
        ("02", "단골이 떠나기 전 어떤 신호도 알 수 없다",
         "방문 주기가 늘어나도, 객단가가 떨어져도 사장님은 모른다. 다른 가게로 옮긴 후에야 알아챈다."),
        ("03", "단골을 관리할 여유가 없다",
         "아침부터 저녁까지 장사 준비·손질·응대·마감까지. 체계적으로 고객을 관리할 시간적·비용적 여유가 부족."),
    ]
    for i, (num, t, b) in enumerate(items):
        x = M_LEFT + (cw + gap) * i
        add_rect(slide, x, cy, cw, ch, fill=WHITE, line=GRAY_300, line_w=1.0, corner=True)
        add_rect(slide, x + Inches(0.3), cy + Inches(0.3), Inches(0.55), Inches(0.55),
                 fill=AWS_ORANGE, line=None, corner=True)
        _, tf = add_text(slide, x + Inches(0.3), cy + Inches(0.3), Inches(0.55), Inches(0.55),
                         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, num, size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                 line_spacing=1.0, first=True)
        _, tf = add_text(slide, x + Inches(0.3), cy + Inches(1.05), cw - Inches(0.6), Inches(0.8))
        add_para(tf, t, size=15, bold=True, color=GRAY_900, line_spacing=1.25, first=True)
        _, tf = add_text(slide, x + Inches(0.3), cy + Inches(1.9), cw - Inches(0.6), ch - Inches(2.1))
        add_para(tf, b, size=12, color=GRAY_700, line_spacing=1.4, first=True)


# ---- 5. Demand hypothesis ----
def s05(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 5, "“정육점에서 고기 사는데 CRM이 필요한가?” — 정면으로 답합니다.",
               eyebrow="DEMAND HYPOTHESIS",
               subtitle="없는 수요를 만들려는 게 아니라, 이미 존재하지만 도구가 없는 수요를 봅니다.")
    # Left: evidence table
    lx = M_LEFT
    lw = Inches(7.4)
    _, tf = add_text(slide, lx, Inches(2.35), lw, Inches(0.3))
    add_para(tf, "정황 근거 (Why we believe)", size=12, bold=True,
             color=AWS_ORANGE_DARK, line_spacing=1.0, first=True)
    rows = [
        ("①", "정육점 매출의 70%가 단골 반복구매 (파일럿)", "단골 = 매출의 본체"),
        ("②", "명절·기념일 매출이 평소의 3~5배", "타이밍 마케팅 여지 큼"),
        ("③", "사장님은 이미 단골 챙기려 시도 (전화·외상장부)", "잠재 수요 有, 도구 無"),
        ("④", "“단골 알면 매출 다른데 다 못 외운다” (인터뷰)", "의지 有, 물리적 불가"),
    ]
    table(slide, lx, Inches(2.7), lw,
          [Inches(0.45), Inches(4.55), Inches(2.4)], rows,
          header=["", "근거", "시사점"], row_h=Inches(0.72), font=10.5, header_font=10,
          aligns=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT])

    # Right: honest confession + promise
    rx = M_LEFT + lw + Inches(0.3)
    rw = CONTENT_W - lw - Inches(0.3)
    add_rect(slide, rx, Inches(2.35), rw, Inches(4.15),
             fill=AWS_DARK_NAVY, line=None, corner=True)
    _, tf = add_text(slide, rx + Inches(0.3), Inches(2.6), rw - Inches(0.6), Inches(3.7))
    add_para(tf, "정직한 고백 — 그리고 약속", size=14, bold=True, color=AWS_ORANGE_LIGHT,
             line_spacing=1.2, first=True)
    add_para(tf, "아직 정량 증거(도입 전/후 매출)는 없습니다. 가장 자주 받는 지적이고, 우리도 압니다.",
             size=11, color=WHITE, line_spacing=1.4, space_before=8)
    add_para(tf, "그래서 이 Pre-Seed는", size=11, color=WHITE,
             line_spacing=1.4, space_before=8)
    add_para(tf, "‘제품을 만드는 돈’이 아니라\n‘수요를 숫자로 증명하는 돈’.",
             size=14, bold=True, color=AWS_ORANGE_LIGHT, line_spacing=1.35)
    add_para(tf, "Stage 1(6개월) 단일 미션: 파일럿 매장 도입 전/후 매출로 수요 증명. 실패 시 제품 PIVOT.",
             size=11, color=WHITE, line_spacing=1.4, space_before=8)
    _, tf = add_text(slide, rx + Inches(0.3), Inches(6.05), rw - Inches(0.6), Inches(0.4))
    add_para(tf, "“6개월 뒤, 가설이 아니라 숫자를 들고 옵니다.”",
             size=10.5, italic=True, color=AWS_ORANGE_LIGHT, line_spacing=1.2, first=True)


# ---- 6. Solution ----
def s06(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 6, "사장님은 눈앞의 손님에게 집중하세요. 나머지는 또와가.",
               eyebrow="SOLUTION")
    lx = M_LEFT
    lw = Inches(4.4)
    _, tf = add_text(slide, lx, Inches(2.3), lw, Inches(0.35))
    add_para(tf, "효과 (제품 가치 가설 — Stage 1 실증 예정)", size=11, bold=True,
             color=AWS_ORANGE_DARK, line_spacing=1.0, first=True)
    mtop = Inches(2.7)
    mh = Inches(1.05)
    metrics = [("주 5초", "시간 절약 — 단골관리에 쓰는 시간"),
               ("−50%", "비용 감소 — 마케팅·재고 비용"),
               ("+15%", "매출 증가 — 매장 매출")]
    for i, (v, d) in enumerate(metrics):
        y = mtop + (mh + Inches(0.15)) * i
        add_rect(slide, lx, y, lw, mh, fill=AWS_ORANGE_FAINT,
                 line=AWS_ORANGE_LIGHT, line_w=1.0, corner=True)
        _, tf = add_text(slide, lx + Inches(0.3), y, Inches(1.6), mh, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, v, size=24, bold=True, color=AWS_ORANGE_DARK, line_spacing=1.0, first=True)
        _, tf = add_text(slide, lx + Inches(1.95), y, lw - Inches(2.15), mh, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, d, size=11.5, color=GRAY_800, line_spacing=1.3, first=True)

    rx = M_LEFT + lw + Inches(0.3)
    rw = CONTENT_W - lw - Inches(0.3)
    _, tf = add_text(slide, rx, Inches(2.3), rw, Inches(0.35))
    add_para(tf, "기능 — 3축", size=11, bold=True, color=AWS_ORANGE_DARK,
             line_spacing=1.0, first=True)
    funcs = [("①", "매장 데이터 기반 단골 관리 (Wedge)",
              "POS 거래 → 단골 자동 분류·메시지·발송·재방문 추적"),
             ("②", "재고량 분석 CRM 활동",
              "과재고 부위 자동 마케팅 · 신선도 기반 추천"),
             ("③", "판매량 예측 기반 발주 제안·자동화",
              "단골 패턴 + 명절·계절 변수 → 발주량 자동")]
    for i, (n, t, b) in enumerate(funcs):
        y = mtop + (mh + Inches(0.15)) * i
        add_rect(slide, rx, y, rw, mh, fill=WHITE, line=GRAY_300, line_w=1.0, corner=True)
        _, tf = add_text(slide, rx + Inches(0.2), y, Inches(0.5), mh,
                         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, n, size=22, bold=True, color=AWS_ORANGE, align=PP_ALIGN.CENTER,
                 line_spacing=1.0, first=True)
        _, tf = add_text(slide, rx + Inches(0.75), y + Inches(0.13), rw - Inches(0.9), Inches(0.4))
        add_para(tf, t, size=12.5, bold=True, color=GRAY_900, line_spacing=1.15, first=True)
        _, tf = add_text(slide, rx + Inches(0.75), y + Inches(0.55), rw - Inches(0.9), Inches(0.45))
        add_para(tf, b, size=10.5, color=GRAY_700, line_spacing=1.3, first=True)

    _, tf = add_text(slide, M_LEFT, Inches(6.5), CONTENT_W, Inches(0.35))
    add_para(tf, "* 효과 수치는 제품 가치 가설 — 실증 데이터 없음, 베타 30매장에서 측정 예정.",
             size=10, italic=True, color=GRAY_500, line_spacing=1.2, first=True)


# ---- 7. Concept image ----
def s07(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 7, "또와가 일하는 모습 (제품 컨셉)",
               eyebrow="PRODUCT CONCEPT",
               subtitle="사장님은 눈앞의 손님에게만 집중. 또와가 동시에 단골·재고·이탈을 챙깁니다.")
    img_x, img_y, img_w, img_h = M_LEFT, Inches(2.4), Inches(6.5), Inches(4.3)
    add_rect(slide, img_x, img_y, img_w, img_h, fill=AWS_ORANGE_FAINT,
             line=AWS_ORANGE_LIGHT, line_w=1.5, corner=True)
    _, tf = add_text(slide, img_x, img_y, img_w, img_h,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "[ 컨셉 이미지 영역 ]", size=16, bold=True, color=AWS_ORANGE_DARK,
             align=PP_ALIGN.CENTER, line_spacing=1.3, first=True)
    add_para(tf, "사장님이 손님 응대 중 + 반투명 고스트 3마리가 동시에 일하는 컷",
             size=12, color=GRAY_700, align=PP_ALIGN.CENTER, line_spacing=1.4)
    add_para(tf, "spec: assets/slide5_concept_image_prompt.md",
             size=9, italic=True, color=GRAY_500, align=PP_ALIGN.CENTER, line_spacing=1.3)
    bx = M_LEFT + img_w + Inches(0.3)
    bw = CONTENT_W - img_w - Inches(0.3)
    btop = Inches(2.4)
    bh = Inches(1.35)
    bubbles = [("A", "정육왕 고객님! 보낸 문자 보고 찾아주셨군요!", "단골관리 · 재방문 유도"),
               ("B", "사장님, 목살 재고가 많아요! 지금 프로모션 돌릴게요!", "재고 분석 · 자동 마케팅"),
               ("C", "오늘 기준 이탈 고객 N명 예상돼요. 쿠폰 보낼까요?", "이탈 예측 · 자동 대응")]
    for i, (lab, msg, tag) in enumerate(bubbles):
        y = btop + (bh + Inches(0.15)) * i
        add_rect(slide, bx, y, bw, bh, fill=WHITE, line=AWS_ORANGE, line_w=1.5, corner=True)
        add_rect(slide, bx + Inches(0.2), y + Inches(0.22), Inches(0.5), Inches(0.5),
                 fill=AWS_ORANGE, line=None, corner=True)
        _, tf = add_text(slide, bx + Inches(0.2), y + Inches(0.22), Inches(0.5), Inches(0.5),
                         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, lab, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                 line_spacing=1.0, first=True)
        _, tf = add_text(slide, bx + Inches(0.85), y + Inches(0.12), bw - Inches(1.0), Inches(0.85))
        add_para(tf, f"“{msg}”", size=11, italic=True, color=GRAY_900, line_spacing=1.3, first=True)
        _, tf = add_text(slide, bx + Inches(0.85), y + Inches(0.95), bw - Inches(1.0), Inches(0.3))
        add_para(tf, f"→ {tag}", size=10, bold=True, color=AWS_ORANGE_DARK, line_spacing=1.2, first=True)


# ---- 8. Functions + consent/data ----
def s08(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 8, "주요 기능 & 데이터 수집·동의 메커니즘",
               eyebrow="FEATURES & DATA")
    # Left: features
    lx = M_LEFT
    lw = Inches(5.3)
    _, tf = add_text(slide, lx, Inches(2.0), lw, Inches(0.35))
    add_para(tf, "A. 주요 기능 — 고객에게 제공하는 것", size=12, bold=True,
             color=AWS_ORANGE_DARK, line_spacing=1.0, first=True)
    feat = [("고객관리", "이탈 예측 · 고객 분석 · 개인화 마케팅 · 재방문 유도"),
            ("판매 관리", "매출 분석 · 마진 분석"),
            ("재고 관리", "재고 예측 · 자동 발주")]
    fy = Inches(2.45)
    fh = Inches(1.05)
    for i, (cat, fns) in enumerate(feat):
        y = fy + (fh + Inches(0.12)) * i
        add_rect(slide, lx, y, lw, fh, fill=WHITE, line=GRAY_300, line_w=1.0, corner=True)
        add_rect(slide, lx, y, Inches(0.08), fh, fill=AWS_ORANGE, line=None)
        _, tf = add_text(slide, lx + Inches(0.25), y + Inches(0.15), lw - Inches(0.5), Inches(0.4))
        add_para(tf, cat, size=14, bold=True, color=GRAY_900, line_spacing=1.1, first=True)
        _, tf = add_text(slide, lx + Inches(0.25), y + Inches(0.55), lw - Inches(0.5), Inches(0.45))
        add_para(tf, fns, size=11, color=GRAY_700, line_spacing=1.3, first=True)

    # Right: consent + data
    rx = M_LEFT + lw + Inches(0.3)
    rw = CONTENT_W - lw - Inches(0.3)
    add_rect(slide, rx, Inches(2.0), rw, Inches(4.6),
             fill=AWS_ORANGE_FAINT, line=AWS_ORANGE_LIGHT, line_w=1.0, corner=True)
    _, tf = add_text(slide, rx + Inches(0.3), Inches(2.2), rw - Inches(0.6), Inches(0.4))
    add_para(tf, "B. 데이터를 어떻게 모으나 — 동의 & 수집", size=12, bold=True,
             color=AWS_ORANGE_DARK, line_spacing=1.1, first=True)
    _, tf = add_text(slide, rx + Inches(0.3), Inches(2.6), rw - Inches(0.6), Inches(0.35))
    add_para(tf, "“고객이 그냥 사가는데 동의를 어떻게?” — 핵심 운영 질문에 답합니다.",
             size=10, italic=True, color=GRAY_600, line_spacing=1.2, first=True)
    _, tf = add_text(slide, rx + Inches(0.3), Inches(3.0), rw - Inches(0.6), Inches(1.7))
    add_para(tf, "동의 수집 흐름 (사장님 손 = 0)", size=11, bold=True, color=GRAY_900,
             line_spacing=1.2, first=True)
    add_para(tf, "1. POS 결제 시 단골 적립·멤버십 유도 (적립 흐름에 개인정보 활용 동의 1회 포함)",
             size=10.5, color=GRAY_800, line_spacing=1.35, space_before=3)
    add_para(tf, "2. 동의 고객만 단골 DB 적재 — 옵트인 (개인정보보호법 준수)",
             size=10.5, color=GRAY_800, line_spacing=1.35)
    add_para(tf, "3. 알림톡도 수신 동의 채널만 (정보통신망법 준수)",
             size=10.5, color=GRAY_800, line_spacing=1.35)
    _, tf = add_text(slide, rx + Inches(0.3), Inches(4.85), rw - Inches(0.6), Inches(1.4))
    add_para(tf, "무슨 데이터를, 어떻게", size=11, bold=True, color=GRAY_900,
             line_spacing=1.2, first=True)
    add_para(tf, "· 거래(POS): 부분육·금액·일시 — 자동 수집",
             size=10.5, color=GRAY_800, line_spacing=1.35, space_before=3)
    add_para(tf, "· 고객: 전화번호·재방문 — 동의 고객만",
             size=10.5, color=GRAY_800, line_spacing=1.35)
    add_para(tf, "· 파생: 방문 주기·객단가·선호 부위 — 자동 산출",
             size=10.5, color=GRAY_800, line_spacing=1.35)

    _, tf = add_text(slide, M_LEFT, Inches(6.7), CONTENT_W, Inches(0.35))
    add_para(tf, "→ 데이터 품질 = 동의율 × POS 연동률. Stage 1에서 동의 전환율도 핵심 측정 지표.",
             size=11, bold=True, italic=True, color=AWS_ORANGE_DARK, line_spacing=1.2, first=True)


# ---- 9. Why Now ----
def s09(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 9, "도메인 특화 AI 에이전트의 시대. 소상공인 자리는 아직 비어있습니다.",
               eyebrow="WHY NOW")
    cy = Inches(2.3)
    ch = Inches(3.4)
    gap = Inches(0.25)
    cw = (CONTENT_W - gap * 2) / 3
    cols = [("01", "AI 활용 비용이 1/10 이하로",
             "LLM API 비용이 3년 사이 폭락. 매장 월 몇 만원으로도 풀스택 AI 운영 가능한 첫 시점."),
            ("02", "도메인 특화 AI 에이전트 등장",
             "법률(Harvey), 의료(Hippocratic AI), 영업(11x) — 수직 도메인에서 한 업무를 완전 자동화하는 AI가 글로벌 등장."),
            ("03", "그런데 소상공인 버티컬은 비어있음",
             "전문직·테크엔 도메인 AI가 빠르게 들어왔지만, 동네 자영업은 아직 0개사. 또와가 정육에서 시작.")]
    for i, (n, t, b) in enumerate(cols):
        x = M_LEFT + (cw + gap) * i
        add_rect(slide, x, cy, cw, ch, fill=WHITE, line=GRAY_300, line_w=1.0, corner=True)
        add_rect(slide, x, cy, cw, Inches(0.08), fill=AWS_ORANGE, line=None)
        _, tf = add_text(slide, x + Inches(0.3), cy + Inches(0.3), cw - Inches(0.6), Inches(0.5))
        add_para(tf, n, size=22, bold=True, color=AWS_ORANGE_DARK, line_spacing=1.0, first=True)
        _, tf = add_text(slide, x + Inches(0.3), cy + Inches(0.95), cw - Inches(0.6), Inches(0.8))
        add_para(tf, t, size=15, bold=True, color=GRAY_900, line_spacing=1.25, first=True)
        _, tf = add_text(slide, x + Inches(0.3), cy + Inches(1.85), cw - Inches(0.6), ch - Inches(2.05))
        add_para(tf, b, size=11.5, color=GRAY_700, line_spacing=1.5, first=True)
    cy2 = cy + ch + Inches(0.2)
    add_rect(slide, M_LEFT, cy2, CONTENT_W, Inches(0.55), fill=AWS_ORANGE_FAINT,
             line=AWS_ORANGE_LIGHT, line_w=1.0, corner=True)
    _, tf = add_text(slide, M_LEFT, cy2, CONTENT_W, Inches(0.55),
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "매크로 순풍 — 한국 1인당 정육 소비 54.6→62.9kg (5년 +15%)  ·  정부 소상공인 디지털 전환 바우처 연 수천억원",
             size=11, italic=True, color=GRAY_700, align=PP_ALIGN.CENTER, line_spacing=1.2, first=True)


# ---- 10. Market ----
def s10(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 10, "정육 → 식품 소매·외식 → 재방문 B2C 소상공인 전체.",
               eyebrow="MARKET")
    by = Inches(2.2)
    bh = Inches(1.5)
    gap = Inches(0.3)
    bw = (CONTENT_W - gap * 2) / 3
    bars = [("TAM", "재방문 기반 B2C 소상공인", "약 300만 업체 / 7조원", AWS_ORANGE_FAINT, AWS_ORANGE_LIGHT, GRAY_900),
            ("SAM", "식품 관련 소매·외식", "약 90만 개 / 1조원", AWS_ORANGE_LIGHT, AWS_ORANGE_MID, GRAY_900),
            ("SOM", "정육점 (첫 진입 시장)", "약 57,000개소 / 2,200억원", AWS_ORANGE, AWS_ORANGE_DARK, WHITE)]
    for i, (lbl, desc, sz, fill, line, tc) in enumerate(bars):
        x = M_LEFT + (bw + gap) * i
        sub = WHITE if tc == WHITE else GRAY_700
        add_rect(slide, x, by, bw, bh, fill=fill, line=line, line_w=1.5, corner=True)
        _, tf = add_text(slide, x + Inches(0.3), by + Inches(0.2), bw - Inches(0.6), Inches(0.5))
        add_para(tf, lbl, size=20, bold=True, color=tc, line_spacing=1.0, first=True)
        _, tf = add_text(slide, x + Inches(0.3), by + Inches(0.7), bw - Inches(0.6), Inches(0.35))
        add_para(tf, desc, size=11, color=sub, line_spacing=1.2, first=True)
        _, tf = add_text(slide, x + Inches(0.3), by + Inches(1.0), bw - Inches(0.6), Inches(0.45))
        add_para(tf, sz, size=14, bold=True, color=tc, line_spacing=1.2, first=True)
    _, tf = add_text(slide, M_LEFT, Inches(3.95), CONTENT_W, Inches(0.3))
    add_para(tf, "Capturable SOM — 시간별 점유율 (벤치마크 기반)", size=12, bold=True,
             color=AWS_ORANGE_DARK, line_spacing=1.0, first=True)
    rows = [("18m (Series A)", "3%", "약 1,700매장", "Pre-Series A SaaS 침투율"),
            ("5년 (안착)", "10%", "약 5,700매장", "Shopify · Toast 수준"),
            ("10년 (리더)", "20%", "약 11,400매장", "Square · Mindbody 수준")]
    table(slide, M_LEFT, Inches(4.3), CONTENT_W,
          [Inches(2.6), Inches(2.0), Inches(3.0), CONTENT_W - Inches(7.6)], rows,
          header=["시점", "점유율", "매장 수", "글로벌 벤치마크"], row_h=Inches(0.5), font=11,
          hl_row=2)
    _, tf = add_text(slide, M_LEFT, Inches(6.55), CONTENT_W, Inches(0.3))
    add_para(tf, "* Square이 미국 SMB POS에서 했던 일을, 또와는 한국 정육 카테고리에서 — Wedge → 운영 OS → 유통.",
             size=10, italic=True, color=GRAY_500, line_spacing=1.2, first=True)


# ---- 11. Business Model — staged + ROI ----
def s11(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 11, "CRM으로 침투해 유통으로 — 단계적으로 둘 다, 단 순서가 명확합니다.",
               eyebrow="BUSINESS MODEL — STAGED")
    # Left: staged strategy table
    lx = M_LEFT
    lw = Inches(6.8)
    _, tf = add_text(slide, lx, Inches(2.0), lw, Inches(0.3))
    add_para(tf, "단계적 전략 — 지금 라운드가 어디인지 명확히", size=11, bold=True,
             color=AWS_ORANGE_DARK, line_spacing=1.0, first=True)
    rows = [("정체성", "정육 마케팅·CRM SaaS", "정육 유통 중개 플랫폼"),
            ("수익", "구독 + 문자 + 마케팅 효과", "유통 매입·중개 수수료"),
            ("파는 것", "소프트웨어만 (HW X)", "유통·물류 (검증 후)"),
            ("검증 미션", "수요·공헌이익 양수", "유통 GMV·중개 마진")]
    table(slide, lx, Inches(2.35), lw,
          [Inches(1.5), Inches(2.9), Inches(2.4)], rows,
          header=["", "1단계 (지금 — Pre-Seed)", "2단계 (Series A 후)"],
          row_h=Inches(0.5), font=10.5, header_font=10)
    _, tf = add_text(slide, lx, Inches(4.75), lw, Inches(0.7))
    add_para(tf, "→ 지금은 1단계. 저울·납품·물류는 수요 증명 후 2단계에서. 지금 팀이 감당 못 할 일을 약속하지 않습니다.",
             size=10.5, bold=True, italic=True, color=AWS_ORANGE_DARK, line_spacing=1.3, first=True)

    # Left-bottom: revenue structure
    _, tf = add_text(slide, lx, Inches(5.5), lw, Inches(0.3))
    add_para(tf, "수익 구조 (1단계)", size=11, bold=True, color=AWS_ORANGE_DARK,
             line_spacing=1.0, first=True)
    rows2 = [("Basic 구독", "₩33,000/월", "GM 75~80%"),
             ("문자 발송", "₩33/건", "LTV/CAC 5배+"),
             ("매장 ARPU", "₩8~12만/월", "월 이탈 3%")]
    table(slide, lx, Inches(5.82), lw,
          [Inches(2.2), Inches(2.3), Inches(2.3)], rows2,
          row_h=Inches(0.32), font=10)

    # Right: 점주 ROI
    rx = M_LEFT + lw + Inches(0.3)
    rw = CONTENT_W - lw - Inches(0.3)
    add_rect(slide, rx, Inches(2.0), rw, Inches(4.7),
             fill=AWS_DARK_NAVY, line=None, corner=True)
    _, tf = add_text(slide, rx + Inches(0.3), Inches(2.2), rw - Inches(0.6), Inches(0.7))
    add_para(tf, "점주 수익 영향", size=14, bold=True, color=AWS_ORANGE_LIGHT,
             line_spacing=1.2, first=True)
    add_para(tf, "구독료 33,000원은 회수되는가? (가정 기반 시뮬레이션)",
             size=10, italic=True, color=WHITE, line_spacing=1.2)
    roi = [("정육 객단가", "약 25,000원"),
           ("정육 마진율", "약 25%"),
           ("회당 마진", "6,250원"),
           ("구독료 회수점", "단골 5~6명/월"),
           ("파일럿 단골 풀", "2,100명")]
    ry = Inches(3.1)
    rh = Inches(0.5)
    for i, (k, v) in enumerate(roi):
        y = ry + rh * i
        col = AWS_ORANGE_LIGHT if i >= 3 else WHITE
        _, tf = add_text(slide, rx + Inches(0.3), y, rw - Inches(2.0), rh, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, k, size=11, color=WHITE, line_spacing=1.0, first=True)
        _, tf = add_text(slide, rx + rw - Inches(2.1), y, Inches(1.8), rh,
                         anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)
        add_para(tf, v, size=12, bold=(i >= 3), color=col, align=PP_ALIGN.RIGHT,
                 line_spacing=1.0, first=True)
        if i < len(roi) - 1:
            add_line(slide, rx + Inches(0.3), y + rh, rx + rw - Inches(0.3), y + rh,
                     color=GRAY_700, weight=0.5)
    _, tf = add_text(slide, rx + Inches(0.3), Inches(5.7), rw - Inches(0.6), Inches(0.9))
    add_para(tf, "단골 5~6명 재방문만 추가돼도 구독료 회수. 그 이상은 전부 점주 순증. (Stage 1 실측)",
             size=10.5, bold=True, color=AWS_ORANGE_LIGHT, line_spacing=1.35, first=True)


# ---- 12. Traction ----
def s12(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 12, "파일럿 1매장에서 Zero-Task 워크플로 가동 중.",
               eyebrow="TRACTION — 수요 검증 라운드")
    ny = Inches(2.0)
    nh = Inches(1.4)
    gap = Inches(0.3)
    nw = (CONTENT_W - gap * 2) / 3
    nums = [("약 5만 건", "누적 거래"), ("약 2,100명", "누적 단골 (5회+)"), ("1매장", "파일럿 운영 중")]
    for i, (b, l) in enumerate(nums):
        x = M_LEFT + (nw + gap) * i
        add_rect(slide, x, ny, nw, nh, fill=AWS_ORANGE_FAINT, line=AWS_ORANGE_LIGHT, line_w=1.0, corner=True)
        _, tf = add_text(slide, x, ny + Inches(0.2), nw, Inches(0.7),
                         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, b, size=28, bold=True, color=AWS_ORANGE_DARK, align=PP_ALIGN.CENTER,
                 line_spacing=1.0, first=True)
        _, tf = add_text(slide, x, ny + Inches(0.9), nw, Inches(0.4),
                         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, l, size=12, color=GRAY_700, align=PP_ALIGN.CENTER, line_spacing=1.1, first=True)

    # honest box
    hy = Inches(3.6)
    add_rect(slide, M_LEFT, hy, Inches(6.0), Inches(2.0), fill=WHITE, line=AMBER, line_w=1.5, corner=True)
    _, tf = add_text(slide, M_LEFT + Inches(0.3), hy + Inches(0.15), Inches(5.4), Inches(0.4))
    add_para(tf, "[!] 아직 측정되지 않은 것 — 정직 공개", size=12, bold=True, color=AMBER,
             line_spacing=1.0, first=True)
    _, tf = add_text(slide, M_LEFT + Inches(0.3), hy + Inches(0.6), Inches(5.4), Inches(1.3))
    for t in ["· 도입 전/후 매장 매출 차이", "· 알림 → 재방문 정량 효과",
              "· 잔존율·재방문 주기 단축", "· +15% / −50% 실증"]:
        add_para(tf, t, size=11, color=GRAY_800, line_spacing=1.4, first=(t.startswith("· 도입")))

    # reframe box
    rx = M_LEFT + Inches(6.3)
    rw = CONTENT_W - Inches(6.3)
    add_rect(slide, rx, hy, rw, Inches(2.0), fill=AWS_DARK_NAVY, line=None, corner=True)
    _, tf = add_text(slide, rx + Inches(0.3), hy + Inches(0.2), rw - Inches(0.6), Inches(1.7))
    add_para(tf, "이 Pre-Seed의 본질 = 수요 검증 라운드", size=13, bold=True,
             color=AWS_ORANGE_LIGHT, line_spacing=1.25, first=True)
    add_para(tf, "“제품을 더 만들겠다”가 아니라 “위 빈칸을 숫자로 채우겠다.”",
             size=11, color=WHITE, line_spacing=1.4, space_before=6)
    add_para(tf, "제품은 이미 돌아갑니다(파일럿). 증명 안 된 건 수요뿐 — 그게 정확히 이 돈의 목적.",
             size=11, color=WHITE, line_spacing=1.4, space_before=4)

    _, tf = add_text(slide, M_LEFT, Inches(5.85), CONTENT_W, Inches(0.4))
    add_para(tf, "협력사: ㈜더담우 (연 매출 1,000억 정육 유통사) · 5개 점포 영업망 협력 합의 (도입 시작 전)",
             size=11, color=GRAY_700, line_spacing=1.2, first=True)


# ---- 13. Validation ----
def s13(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 13, "Pre-Seed 5억 = 가설 검증 → PMF → 공헌이익 양수화. Series A는 그 다음.",
               eyebrow="VALIDATION TO-DO")
    sy = Inches(2.3)
    sh = Inches(4.1)
    gap = Inches(0.2)
    sw = (CONTENT_W - gap * 3) / 4
    stages = [
        ("Stage 1", "가설 검증", "+0~6m",
         "Q1 도입 전/후 매출\nQ2 알림→재방문 전환\nQ3 이탈군 잔존율",
         "통과 → Stage 2\n실패 → 제품 PIVOT", AWS_ORANGE),
        ("Stage 2", "PMF 찾기", "+6~12m",
         "Q4 N=30 효과 재현\nQ5 유료 전환율\nQ6 문자 발송량 분포",
         "통과 → Stage 3\n실패 → 가격·세그먼트\n     재설계", AWS_ORANGE_DARK),
        ("Stage 3", "공헌이익 양수화", "+12~18m",
         "Q7 더담우 영업 효율\nQ8 직판 CAC\nQ9 POS 수락율\nQ10 ARPU\nQ11 이탈·LTV",
         "통과 → Series A 도전\n실패 → 브릿지 라운드", AWS_DARK_NAVY),
        ("Stage 4", "시장 확장", "Series A 후",
         "Q12 유통 중개 +\n인접 카테고리 PoC",
         "(Pre-Seed 범위 외)", GRAY_500),
    ]
    for i, (lb, nm, tm, qs, gt, col) in enumerate(stages):
        x = M_LEFT + (sw + gap) * i
        add_rect(slide, x, sy, sw, Inches(0.1), fill=col, line=None)
        add_rect(slide, x, sy + Inches(0.1), sw, sh - Inches(0.1), fill=WHITE, line=GRAY_300, line_w=1.0, corner=True)
        _, tf = add_text(slide, x + Inches(0.2), sy + Inches(0.25), sw - Inches(0.4), Inches(0.35))
        add_para(tf, lb, size=11, bold=True, color=col, line_spacing=1.0, first=True)
        _, tf = add_text(slide, x + Inches(0.2), sy + Inches(0.62), sw - Inches(0.4), Inches(0.45))
        add_para(tf, nm, size=15, bold=True, color=GRAY_900, line_spacing=1.1, first=True)
        _, tf = add_text(slide, x + Inches(0.2), sy + Inches(1.12), sw - Inches(0.4), Inches(0.3))
        add_para(tf, tm, size=10, color=GRAY_500, line_spacing=1.0, first=True)
        _, tf = add_text(slide, x + Inches(0.2), sy + Inches(1.5), sw - Inches(0.4), Inches(1.7))
        add_para(tf, qs, size=10, color=GRAY_800, line_spacing=1.4, first=True)
        gy = sy + sh - Inches(0.95)
        add_rect(slide, x + Inches(0.15), gy, sw - Inches(0.3), Inches(0.8),
                 fill=AWS_ORANGE_FAINT, line=AWS_ORANGE_LIGHT, line_w=0.5, corner=True)
        _, tf = add_text(slide, x + Inches(0.25), gy, sw - Inches(0.5), Inches(0.8), anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, gt, size=9, color=GRAY_800, line_spacing=1.3, first=True)
    _, tf = add_text(slide, M_LEFT, Inches(6.65), CONTENT_W, Inches(0.3))
    add_para(tf, "의사결정 게이트 — +6m: Stage 1 → 베타 가속 · +12m: Stage 2 → 유료 GTM · +18m: Stage 3 → Series A",
             size=10, italic=True, color=GRAY_600, line_spacing=1.2, first=True)


# ---- 14. Team ----
def s14(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 14, "B2B SaaS + AI 박사 + 정육 도메인 — 한 팀.",
               eyebrow="TEAM — 왜 이 팀이 1단계를 해내는가")
    py = Inches(2.2)
    ph = Inches(3.3)
    gap = Inches(0.15)
    pw = (CONTENT_W - gap * 4) / 5
    team = [("안동철", "대표이사", "한밭대 컴공", "㈜스낵포 CTO 8년\nB2B SaaS 0→1"),
            ("배현혜", "AI·데이터 총괄", "Columbia 박사", "빅데이터 전임연구원\n㈜팬블러 창업"),
            ("SEO CHARLES", "제품·UX 총괄", "KAIST 학·석사", "㈜팬블러\nCo-Founder/PO"),
            ("이지백", "현장·도메인", "충남대 경영 박사", "정육 유통사\n사외이사"),
            ("박동일", "운영·재무", "한밭대 회계학사", "—")]
    for i, (nm, role, edu, exp) in enumerate(team):
        x = M_LEFT + (pw + gap) * i
        add_rect(slide, x, py, pw, ph, fill=WHITE, line=GRAY_300, line_w=1.0, corner=True)
        ps = Inches(1.0)
        add_rect(slide, x + (pw - ps) / 2, py + Inches(0.3), ps, ps,
                 fill=AWS_ORANGE_LIGHT, line=AWS_ORANGE_LIGHT, corner=True)
        _, tf = add_text(slide, x, py + Inches(1.4), pw, Inches(0.4), align=PP_ALIGN.CENTER)
        add_para(tf, nm, size=13, bold=True, color=GRAY_900, align=PP_ALIGN.CENTER, line_spacing=1.1, first=True)
        _, tf = add_text(slide, x, py + Inches(1.78), pw, Inches(0.32), align=PP_ALIGN.CENTER)
        add_para(tf, role, size=10, bold=True, color=AWS_ORANGE_DARK, align=PP_ALIGN.CENTER, line_spacing=1.1, first=True)
        _, tf = add_text(slide, x + Inches(0.1), py + Inches(2.12), pw - Inches(0.2), Inches(0.35), align=PP_ALIGN.CENTER)
        add_para(tf, edu, size=9, italic=True, color=GRAY_500, align=PP_ALIGN.CENTER, line_spacing=1.1, first=True)
        _, tf = add_text(slide, x + Inches(0.1), py + Inches(2.5), pw - Inches(0.2), Inches(0.75), align=PP_ALIGN.CENTER)
        add_para(tf, exp, size=9, color=GRAY_700, align=PP_ALIGN.CENTER, line_spacing=1.3, first=True)
    fy = Inches(5.7)
    add_rect(slide, M_LEFT, fy, CONTENT_W, Inches(1.1), fill=AWS_ORANGE_FAINT,
             line=AWS_ORANGE_LIGHT, line_w=1.0, corner=True)
    _, tf = add_text(slide, M_LEFT + Inches(0.4), fy + Inches(0.15), CONTENT_W - Inches(0.8), Inches(0.85))
    add_para(tf, "Founder-Market Fit — B2B SaaS 실행력 + AI 박사급 + 정육 도메인 네트워크",
             size=12, bold=True, color=AWS_ORANGE_DARK, line_spacing=1.2, first=True)
    add_para(tf, "1단계(CRM SaaS)는 정확히 이 팀의 코어 역량. 2단계(유통)는 ㈜더담우(연 1,000억)와의 동맹으로 보완 — 하드웨어·물류를 처음부터 다 하지 않습니다.",
             size=10.5, color=GRAY_700, line_spacing=1.3, space_before=4)


# ---- 15. Roadmap ----
def s15(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 15, "1단계: 정육 CRM SaaS → 2단계: 유통 중개 (Series A 후).",
               eyebrow="ROADMAP")
    ty = Inches(2.5)
    line_y = ty + Inches(0.5)
    mks = [("NOW", "파일럿 1", ""), ("+6m", "베타 N=30", "Stage 1"),
           ("+12m", "유료 100~150", "Stage 2"), ("+18m", "공헌이익 양수", "Stage 3"),
           ("Series A 후", "유통 중개+확장", "Stage 4")]
    n = len(mks)
    step = CONTENT_W / (n - 1)
    add_line(slide, M_LEFT, line_y, M_LEFT + CONTENT_W, line_y, color=AWS_ORANGE, weight=2.5)
    for i, (tm, ms, st) in enumerate(mks):
        x = M_LEFT + step * i
        ds = Inches(0.25)
        add_rect(slide, x - ds / 2, line_y - ds / 2, ds, ds, fill=AWS_ORANGE_DARK,
                 line=WHITE, line_w=2.0, corner=True)
        _, tf = add_text(slide, x - Inches(1.0), line_y - Inches(0.55), Inches(2.0), Inches(0.3), align=PP_ALIGN.CENTER)
        add_para(tf, tm, size=11, bold=True, color=AWS_ORANGE_DARK, align=PP_ALIGN.CENTER, line_spacing=1.0, first=True)
        _, tf = add_text(slide, x - Inches(1.2), line_y + Inches(0.2), Inches(2.4), Inches(0.4), align=PP_ALIGN.CENTER)
        add_para(tf, ms, size=12, bold=True, color=GRAY_900, align=PP_ALIGN.CENTER, line_spacing=1.2, first=True)
        if st:
            _, tf = add_text(slide, x - Inches(1.0), line_y + Inches(0.62), Inches(2.0), Inches(0.3), align=PP_ALIGN.CENTER)
            add_para(tf, st, size=10, italic=True, color=GRAY_500, align=PP_ALIGN.CENTER, line_spacing=1.0, first=True)

    # Stage bands
    band_y = Inches(4.3)
    bw1 = Inches(8.8)
    add_rect(slide, M_LEFT, band_y, bw1, Inches(0.45), fill=AWS_ORANGE_LIGHT, line=AWS_ORANGE, line_w=1.0, corner=True)
    _, tf = add_text(slide, M_LEFT + Inches(0.2), band_y, bw1 - Inches(0.4), Inches(0.45), anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "1단계: 정육 CRM SaaS (지금 라운드)", size=11, bold=True, color=AWS_ORANGE_DARK, line_spacing=1.0, first=True)
    sa_x = M_LEFT + bw1 + Inches(0.2)
    sa_w = CONTENT_W - bw1 - Inches(0.2)
    add_rect(slide, sa_x, band_y, sa_w, Inches(0.45), fill=AWS_DARK_NAVY, line=None, corner=True)
    _, tf = add_text(slide, sa_x, band_y, sa_w, Inches(0.45), align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "2단계: 유통 →", size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, line_spacing=1.0, first=True)

    band_y2 = band_y + Inches(0.55)
    add_rect(slide, M_LEFT, band_y2, bw1, Inches(0.45), fill=AWS_ORANGE_FAINT, line=AWS_ORANGE_LIGHT, line_w=1.0, corner=True)
    _, tf = add_text(slide, M_LEFT + Inches(0.2), band_y2, bw1 - Inches(0.4), Inches(0.45), anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "Pre-Seed 5억(희석) + 정부 R&D 7억(비희석, 디딤돌+TIPS+후속)", size=11, bold=True, color=GRAY_800, line_spacing=1.0, first=True)

    # Series A conditions
    cy = Inches(5.7)
    _, tf = add_text(slide, M_LEFT, cy, CONTENT_W, Inches(0.3))
    add_para(tf, "Series A 진입 조건 (+18m)", size=11, bold=True, color=AWS_ORANGE_DARK, line_spacing=1.0, first=True)
    conds = ["공헌이익 양수", "유료 매장 100~150", "LTV/CAC 3~5배", "월 이탈율 <5%", "더담우 + POS사"]
    cw = CONTENT_W / len(conds)
    for i, c in enumerate(conds):
        x = M_LEFT + cw * i
        add_rect(slide, x + Inches(0.05), cy + Inches(0.35), cw - Inches(0.1), Inches(0.45),
                 fill=WHITE, line=GRAY_300, line_w=0.75, corner=True)
        _, tf = add_text(slide, x + Inches(0.05), cy + Inches(0.35), cw - Inches(0.1), Inches(0.45),
                         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_para(tf, c, size=10, color=GRAY_800, align=PP_ALIGN.CENTER, line_spacing=1.1, first=True)


# ---- 16. The Ask ----
def s16(prs):
    slide = slide_blank(prs)
    add_chrome(slide, 16, "Pre-Seed 5억 + R&D 7억 — 18개월, 수요 증명 → 공헌이익 양수까지.",
               eyebrow="THE ASK")
    hy = Inches(2.0)
    hh = Inches(1.5)
    gap = Inches(0.3)
    hw = (CONTENT_W - gap) / 2
    add_rect(slide, M_LEFT, hy, hw, hh, fill=AWS_ORANGE, line=AWS_ORANGE_DARK, line_w=1.5, corner=True)
    _, tf = add_text(slide, M_LEFT + Inches(0.4), hy + Inches(0.2), hw - Inches(0.8), Inches(0.45))
    add_para(tf, "Pre-Seed", size=14, bold=True, color=WHITE, line_spacing=1.0, first=True)
    _, tf = add_text(slide, M_LEFT + Inches(0.4), hy + Inches(0.6), hw - Inches(0.8), Inches(0.7))
    add_para(tf, "5억  (희석)", size=36, bold=True, color=WHITE, line_spacing=1.0, first=True)
    x2 = M_LEFT + hw + gap
    add_rect(slide, x2, hy, hw, hh, fill=AWS_DARK_NAVY, line=None, corner=True)
    _, tf = add_text(slide, x2 + Inches(0.4), hy + Inches(0.2), hw - Inches(0.8), Inches(0.45))
    add_para(tf, "정부 R&D", size=14, bold=True, color=WHITE, line_spacing=1.0, first=True)
    _, tf = add_text(slide, x2 + Inches(0.4), hy + Inches(0.6), hw - Inches(0.8), Inches(0.7))
    add_para(tf, "7억  (비희석)", size=36, bold=True, color=WHITE, line_spacing=1.0, first=True)
    _, tf = add_text(slide, x2 + Inches(0.4), hy + hh - Inches(0.35), hw - Inches(0.8), Inches(0.3))
    add_para(tf, "디딤돌 2.668억 확보 + TIPS·후속 가설", size=10, italic=True, color=AWS_ORANGE_LIGHT, line_spacing=1.0, first=True)

    _, tf = add_text(slide, M_LEFT, Inches(3.7), CONTENT_W, Inches(0.3))
    add_para(tf, "자금 사용 계획 — 총 12억 통합", size=12, bold=True, color=AWS_ORANGE_DARK, line_spacing=1.0, first=True)
    rows = [("인력", "5.0억", "42%", "영업·운영 + R&D·개발"),
            ("마케팅·영업", "3.0억", "25%", "GTM · 정육 협회 · 바우처 채널"),
            ("제품 개발", "3.0억", "25%", "AI 모델 · POS 연동 · SaaS화 (R&D 자금)"),
            ("기타 운영비", "1.0억", "8%", "사무실 · 법무 · 회계"),
            ("합계", "12.0억", "100%", "18개월 Runway")]
    table(slide, M_LEFT, Inches(4.05), CONTENT_W,
          [Inches(2.5), Inches(1.5), Inches(1.5), CONTENT_W - Inches(5.5)], rows,
          header=["항목", "금액", "비율", "용도"], row_h=Inches(0.42), font=10.5, hl_row=4)
    # efficiency narrative
    ey = Inches(6.35)
    add_rect(slide, M_LEFT, ey, CONTENT_W, Inches(0.55), fill=AWS_ORANGE_FAINT,
             line=AWS_ORANGE_LIGHT, line_w=1.0, corner=True)
    _, tf = add_text(slide, M_LEFT + Inches(0.3), ey, CONTENT_W - Inches(0.6), Inches(0.55), anchor=MSO_ANCHOR.MIDDLE)
    add_para(tf, "CRM×SaaS 펀딩 비우호 환경 → 비희석 R&D 7억 + 공헌이익 양수 목표. “돈을 태우는 회사가 아니라 효율로 증명하는 회사.”",
             size=10.5, bold=True, italic=True, color=AWS_ORANGE_DARK, line_spacing=1.2, first=True)


# ---- 17. Vision ----
def s17(prs):
    slide = slide_blank(prs)
    add_rect(slide, 0, 0, Inches(0.3), SLIDE_H, fill=AWS_ORANGE, line=None)
    _, tf = add_text(slide, Inches(1.0), Inches(0.8), Inches(11), Inches(0.4))
    add_para(tf, "VISION", size=12, bold=True, color=AWS_ORANGE, line_spacing=1.0, first=True)
    _, tf = add_text(slide, Inches(1.0), Inches(1.5), Inches(11.5), Inches(3.0))
    add_para(tf, "기술 발전의 혜택에서 소외된", size=29, bold=True, color=GRAY_900, line_spacing=1.4, first=True)
    add_para(tf, "소상공인의 도메인에 깊숙이 파고들어", size=29, bold=True, color=GRAY_900, line_spacing=1.4)
    add_para(tf, "기술과의 거리감을 줄이는 비즈니스를 하겠습니다.", size=29, bold=True, color=AWS_ORANGE_DARK, line_spacing=1.4)
    _, tf = add_text(slide, Inches(1.0), Inches(4.4), Inches(11.5), Inches(0.45))
    add_para(tf, "— 정육점은 그 시작입니다.", size=16, italic=True, color=GRAY_600, line_spacing=1.2, first=True)
    # staged vision
    vy = Inches(5.15)
    vh = Inches(1.1)
    gap = Inches(0.3)
    vw = (Inches(11.5) - gap * 2) / 3
    vis = [("1단계", "정육 Zero-Task 운영 OS", "CRM·마케팅·재고"),
           ("2단계", "정육 유통 중개 플랫폼", "납품·물류 디지털화"),
           ("3단계", "인접 수직 카테고리 확장", "AI 직원 모델")]
    for i, (st, t, d) in enumerate(vis):
        x = Inches(1.0) + (vw + gap) * i
        add_rect(slide, x, vy, vw, vh, fill=AWS_ORANGE_FAINT, line=AWS_ORANGE_LIGHT, line_w=1.0, corner=True)
        _, tf = add_text(slide, x + Inches(0.2), vy + Inches(0.12), vw - Inches(0.4), Inches(0.3))
        add_para(tf, st, size=11, bold=True, color=AWS_ORANGE_DARK, line_spacing=1.0, first=True)
        _, tf = add_text(slide, x + Inches(0.2), vy + Inches(0.42), vw - Inches(0.4), Inches(0.35))
        add_para(tf, t, size=12, bold=True, color=GRAY_900, line_spacing=1.15, first=True)
        _, tf = add_text(slide, x + Inches(0.2), vy + Inches(0.78), vw - Inches(0.4), Inches(0.3))
        add_para(tf, d, size=10, color=GRAY_600, line_spacing=1.1, first=True)
    _, tf = add_text(slide, Inches(1.0), Inches(6.5), Inches(10), Inches(0.4))
    add_para(tf, "㈜위브원  ·  또와  ·  안동철 대표", size=12, color=GRAY_600, line_spacing=1.2, first=True)
    page_num_only(slide, 17)


# ============================================================
def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    for b in [s01, s02, s03, s04, s05, s06, s07, s08, s09,
              s10, s11, s12, s13, s14, s15, s16, s17]:
        b(prs)
    out_dir = Path("/tmp/butcher-build")
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / "또와_IR덱_v2.0_AWS.pptx"
    prs.save(out)
    print(f"[OK] saved → {out}")
    print(f"     slides: {len(prs.slides)} / target {TOTAL}")


if __name__ == "__main__":
    build()
