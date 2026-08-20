# -*- coding: utf-8 -*-
"""Rebuild slides 2-4 of the Klüber EAL grease deck in the simple KLC-SM style.

Slide 1 (Wilhelmsen Greases Solution cover) is carried over untouched; the three
test slides are re-laid out as a navy title band over a plain white body, matching
the KLC-SM reference slide.
"""
import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC = "/root/.claude/uploads/eae68b7d-8c95-56e3-85c5-ff5decc49086/38bb7dfc-Kluber_EAL_Grease.pptx"
MEDIA = "/tmp/claude-0/-home-user-github-test/eae68b7d-8c95-56e3-85c5-ff5decc49086/scratchpad/klu/"
OUT = "/home/user/github-test/Kluber_EAL_Grease_simple.pptx"

NAVY = RGBColor(0x17, 0x2A, 0x4A)
INK = RGBColor(0x1A, 0x1A, 0x1A)
GREY = RGBColor(0x59, 0x59, 0x59)
LIGHT = RGBColor(0xEF, 0xEF, 0xEF)
LINE = RGBColor(0xD0, 0xD4, 0xDA)
ACCENT = RGBColor(0x1F, 0x6F, 0x3D)
RED = RGBColor(0xC0, 0x00, 0x00)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

KF = "맑은 고딕"
SW, SH = Inches(13.333), Inches(7.5)
BAND_H = Inches(1.12)
FOOT = ("Classified as Internal - This content is proprietary information "
        "intended for employees and partners only.")


def kfont(run, size, bold=False, color=INK):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = KF
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {})
            rPr.append(e)
        e.set("typeface", KF)


def textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.04)
    tf.margin_top = tf.margin_bottom = 0
    return tb, tf


def para(tf, text, size, bold=False, color=INK, space_after=4, first=False,
         align=PP_ALIGN.LEFT, line=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after)
    if line:
        p.line_spacing = line
    r = p.add_run()
    r.text = text
    kfont(r, size, bold, color)
    return p


def rect(slide, x, y, w, h, fill, line=None):
    from pptx.enum.shapes import MSO_SHAPE
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(0.75)
    sh.shadow.inherit = False
    return sh


def chip(slide, x, y, w, text):
    """Light grey section label, like the reference slide's date chips."""
    h = Inches(0.36)
    rect(slide, x, y, w, h, LIGHT)
    tb, tf = textbox(slide, x + Inches(0.12), y, w - Inches(0.2), h,
                     anchor=MSO_ANCHOR.MIDDLE)
    para(tf, text, 12, True, RGBColor(0x33, 0x33, 0x33), 0, first=True)
    return y + h + Inches(0.14)


def frame(slide, title):
    rect(slide, 0, 0, SW, BAND_H, NAVY)
    tb, tf = textbox(slide, Inches(0.55), Inches(0.16), Inches(12.2),
                     BAND_H - Inches(0.3), anchor=MSO_ANCHOR.MIDDLE)
    para(tf, title, 25, True, WHITE, 0, first=True)
    fb, ff = textbox(slide, Inches(0.55), SH - Inches(0.42), Inches(12.2), Inches(0.3))
    para(ff, FOOT, 10, False, RGBColor(0x8C, 0x8C, 0x8C), 0, first=True)


def table(slide, x, y, w, rows, col_w, head=True, row_h=Inches(0.34)):
    """Simple bordered table drawn from rectangles + textboxes."""
    cy = y
    for ri, row in enumerate(rows):
        cx = x
        h = row_h
        for ci, cell in enumerate(row):
            cw = col_w[ci]
            is_head = head and ri == 0
            rect(slide, cx, cy, cw, h,
                 NAVY if is_head else WHITE, None if is_head else LINE)
            tb, tf = textbox(slide, cx + Inches(0.08), cy, cw - Inches(0.16), h,
                             anchor=MSO_ANCHOR.MIDDLE)
            txt, bold, col, align = cell if isinstance(cell, tuple) else (cell, False, INK, PP_ALIGN.LEFT)
            if is_head:
                bold, col = True, WHITE
            para(tf, txt, 11.5, bold, col, 0, first=True, align=align)
            cx += cw
        cy += h
    return cy


# ---------------------------------------------------------------- build
prs = Presentation(SRC)

# drop slides 2-4, keep the cover
sldIdLst = prs.slides._sldIdLst
for sldId in list(sldIdLst)[1:]:
    prs.part.drop_rel(sldId.rId)
    sldIdLst.remove(sldId)

blank = prs.slide_layouts[0]


def new_slide(title):
    s = prs.slides.add_slide(blank)
    for shp in list(s.shapes):
        shp._element.getparent().remove(shp._element)
    frame(s, title)
    return s


L = Inches(0.55)          # left margin
COLW = Inches(6.05)       # column width
R = Inches(6.95)          # right column x
TOP = Inches(1.45)

# ---------------------------------------------------------------- slide 2
s = new_slide("Klüberbio AM 92-142  —  적점(Dropping Point) 시험")

y = chip(s, L, TOP, COLW, "시험 개요 —  DIN ISO 2176")
tb, tf = textbox(s, L, y, COLW, Inches(1.6))
para(tf, "그리스가 시험 조건에서 반고체 상태로부터 액체 상태로 전환되는 온도.",
     13, False, INK, 6, first=True, line=1.25)
para(tf, "곧, 그리스가 구조를 유지하며 제 기능을 수행할 수 있는 최고 온도를 뜻합니다.",
     13, False, INK, 10, line=1.25)
para(tf, "※ 무기 증점제(Inorganic thickener)는 용융되지 않습니다.", 11.5, True, GREY, 0)

y2 = y + Inches(1.75)
y2 = chip(s, L, y2, COLW, "증점제 방식별 특성")
table(s, L, y2, COLW,
      [["증점제", "적점", "Zn 부식 / 생분해성"],
       [("무기 증점제 · Complex Soap", False, INK, PP_ALIGN.LEFT), ("높음", False, INK, PP_ALIGN.CENTER), ("제약 있음", False, INK, PP_ALIGN.CENTER)],
       [("Calcium Single Soap", False, INK, PP_ALIGN.LEFT), ("낮음", False, INK, PP_ALIGN.CENTER), ("양호", False, INK, PP_ALIGN.CENTER)],
       [("Hybrid Thickener  (AM 92-142)", True, ACCENT, PP_ALIGN.LEFT), ("높음", True, ACCENT, PP_ALIGN.CENTER), ("양호", True, ACCENT, PP_ALIGN.CENTER)]],
      [Inches(2.95), Inches(1.35), Inches(1.75)])

y = chip(s, R, TOP, COLW, "결과")
rect(s, R, y, COLW, Inches(1.5), LIGHT)
tb, tf = textbox(s, R, y + Inches(0.14), COLW, Inches(1.25), anchor=MSO_ANCHOR.TOP)
para(tf, "> 170 °C", 40, True, NAVY, 2, first=True, align=PP_ALIGN.CENTER)
para(tf, "Dropping Point", 12, False, GREY, 0, align=PP_ALIGN.CENTER)

yy = y + Inches(1.68)
tb, tf = textbox(s, R, yy, COLW, Inches(0.9))
para(tf, "Solution : Hybrid Thickener", 15, True, INK, 4, first=True)
para(tf, "Calcium single soap 기반 하이브리드 증점제 — 높은 적점과 생분해성을 함께 확보",
     12.5, False, GREY, 0, line=1.2)

s.shapes.add_picture(MEDIA + "image7.jpeg", R + Inches(2.15), yy + Inches(1.0),
                     height=Inches(1.85))
tb, tf = textbox(s, R, yy + Inches(2.95), COLW, Inches(0.3))
para(tf, "동적 와이어로프 운용 시 고온에서도 피막 유지", 12, True, INK, 0,
     first=True, align=PP_ALIGN.CENTER)

# ---------------------------------------------------------------- slide 3
s = new_slide("Klüberbio AM 92-142  —  부착력 / Throw-off Resistance 시험")

y = chip(s, L, TOP, COLW, "시험 조건 —  회전 2,000 rpm · 120초 · 70 °C")
s.shapes.add_picture(MEDIA + "image8.png", L + Inches(0.6), y, width=Inches(4.85))
tb, tf = textbox(s, L, y + Inches(2.05), COLW, Inches(0.3))
para(tf, "Disk specimen 에 도포된 그리스의 비산 저항 측정", 12, False, GREY, 0,
     first=True, align=PP_ALIGN.CENTER)

y2 = y + Inches(2.5)
y2 = chip(s, L, y2, COLW, "잔존 그리스량 (Residual grease weight, mg/cm²)")
end = table(s, L, y2, COLW,
            [["제품", "잔존량", "비교"],
             [("Klüberbio AM 92-142  (EAL)", True, ACCENT, PP_ALIGN.LEFT), ("12.2", True, ACCENT, PP_ALIGN.CENTER), ("약 10배", True, ACCENT, PP_ALIGN.CENTER)],
             [("Competitor grease (Mineral oil)", False, INK, PP_ALIGN.LEFT), ("1.05", False, INK, PP_ALIGN.CENTER), ("기준", False, GREY, PP_ALIGN.CENTER)]],
            [Inches(3.05), Inches(1.5), Inches(1.5)])
tb, tf = textbox(s, L, end + Inches(0.12), COLW, Inches(0.3))
para(tf, "높은 잔존 그리스량 → 로프 보호력 우수", 13, True, INK, 0, first=True)

y = chip(s, R, TOP, COLW, "비산 시 위험")
pw = Inches(2.9)
s.shapes.add_picture(MEDIA + "image9.png", R, y, width=pw)
s.shapes.add_picture(MEDIA + "image10.png", R + Inches(3.15), y, width=pw)

cy = y + Inches(1.95)
for cx, txt in ((R, "조종실 전면 유리에 그리스 낙하\n시야 방해 → 안전 위험"),
                (R + Inches(3.15), "브레이크 디스크로 그리스 비산\n제동 성능 저하 우려")):
    tb, tf = textbox(s, cx, cy, pw, Inches(0.75))
    for i, ln in enumerate(txt.split("\n")):
        para(tf, ln, 12, i == 1, RED if i == 1 else INK, 2, first=(i == 0),
             align=PP_ALIGN.CENTER)

y3 = cy + Inches(1.0)
rect(s, R, y3, COLW, Inches(0.95), LIGHT)
tb, tf = textbox(s, R + Inches(0.2), y3, COLW - Inches(0.4), Inches(0.95),
                 anchor=MSO_ANCHOR.MIDDLE)
para(tf, "비부착성 그리스는 로프를 보호하지 못할 뿐 아니라, "
         "주변 설비로 튀어 안전 문제로 이어집니다.", 12.5, False, INK, 0,
     first=True, line=1.25)

# ---------------------------------------------------------------- slide 4
s = new_slide("Klüberbio AM 92-142  —  −60 °C 굴곡 스트립(Bending Strip) 시험")

y = chip(s, L, TOP, COLW, "시험 목적")
tb, tf = textbox(s, L, y, COLW, Inches(0.9))
para(tf, "저온에서 그리스는 건조 · 경화되는 경향이 있고,", 13, False, INK, 4,
     first=True, line=1.25)
para(tf, "이로 인해 윤활 피막이 부착력을 잃게 됩니다.", 13, False, INK, 0, line=1.25)

y2 = y + Inches(1.05)
y2 = chip(s, L, y2, COLW, "시험 절차")
steps = ["두께 1 mm 윤활 피막을 얇은 금속판에 코팅",
         "원하는 저온으로 24시간 냉각",
         "시험 샘플을 실린더에 감아 굴곡 부여"]
for i, st in enumerate(steps):
    ry = y2 + Inches(0.52) * i
    rect(s, L, ry, Inches(0.42), Inches(0.42), NAVY)
    tb, tf = textbox(s, L, ry, Inches(0.42), Inches(0.42), anchor=MSO_ANCHOR.MIDDLE)
    para(tf, str(i + 1), 13, True, WHITE, 0, first=True, align=PP_ALIGN.CENTER)
    tb, tf = textbox(s, L + Inches(0.6), ry, COLW - Inches(0.6), Inches(0.42),
                     anchor=MSO_ANCHOR.MIDDLE)
    para(tf, st, 12.5, False, INK, 0, first=True)

y3 = y2 + Inches(1.75)
rect(s, L, y3, COLW, Inches(0.85), LIGHT)
tb, tf = textbox(s, L + Inches(0.2), y3, COLW - Inches(0.4), Inches(0.85),
                 anchor=MSO_ANCHOR.MIDDLE)
para(tf, "피막이 손상되거나 벗겨지지 않는 한, 시험 온도까지 윤활 기능을 유지합니다.",
     12.5, False, INK, 0, first=True, line=1.25)

y4 = y3 + Inches(1.05)
tb, tf = textbox(s, L, y4, COLW, Inches(0.4))
para(tf, "유동점(Pour Point)  −30 °C  —  저온 환경 운용에 유리", 13.5, True, ACCENT, 0,
     first=True)

y = chip(s, R, TOP, COLW, "−60 °C 저온 굴곡 시험 결과")
pw = Inches(2.9)                       # 967x514 -> h 1.54"
s.shapes.add_picture(MEDIA + "image12.jpeg", R, y, width=pw)
s.shapes.add_picture(MEDIA + "image13.jpeg", R + Inches(3.15), y, width=pw)
cy = y + Inches(1.66)
for cx, txt in ((R, "코팅 금속판\n균열 · 박리 없음"),
                (R + Inches(3.15), "실린더 굴곡\n부착 상태 유지")):
    tb, tf = textbox(s, cx, cy, pw, Inches(0.75))
    for i, ln in enumerate(txt.split("\n")):
        para(tf, ln, 12, i == 1, INK if i == 1 else GREY, 2, first=(i == 0),
             align=PP_ALIGN.CENTER)

y5 = cy + Inches(0.95)
rect(s, R, y5, COLW, Inches(1.0), LIGHT)
tb, tf = textbox(s, R + Inches(0.2), y5, COLW - Inches(0.4), Inches(1.0),
                 anchor=MSO_ANCHOR.MIDDLE)
para(tf, "−60 °C 에서 24시간 냉각 후 굴곡을 주어도 피막이 갈라지거나 "
         "벗겨지지 않았습니다.", 12.5, False, INK, 0, first=True, line=1.25)

prs.save(OUT)
print("saved", OUT, "| slides:", len(prs.slides.__iter__.__self__._sldIdLst))
