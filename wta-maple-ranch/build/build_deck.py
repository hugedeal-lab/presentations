"""Build the WTA Maple Ranch training deck "Mastering Agency Commercials" on the WTA 2026 master template.

Usage: python3 build_deck.py [output.pptx]

Every slide sits on a native WTA layout (backgrounds, logo, footer rule, slide number),
and all type uses the template's DM Sans families and theme palette.
"""
import copy
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.shapes import PP_PLACEHOLDER as PH
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE.parent / "template" / "WTA-2026-MasterSlideDeck.pptx"
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE.parent / "WTA-Mastering-Agency-Commercials-Maple-Ranch.pptx"

# --- WTA theme palette (from the template's "WTA Colors" scheme) ---------------------
NAVY = RGBColor(0x12, 0x2A, 0x4E)
BLUE = RGBColor(0x00, 0x50, 0x9A)
RED = RGBColor(0xFF, 0x33, 0x33)
ORANGE = RGBColor(0xFF, 0x66, 0x33)
TEAL = RGBColor(0x66, 0xCC, 0xCC)
CYAN = RGBColor(0x00, 0xA3, 0xDA)
TEXT = RGBColor(0x33, 0x33, 0x33)
MUTED = RGBColor(0x6B, 0x6B, 0x6B)
GREY = RGBColor(0xF1, 0xF1, 0xF1)
RULE = RGBColor(0xC8, 0xC8, 0xC8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

BODY = "DM Sans 14pt"
LIGHT = "DM Sans 14pt ExtraLight"
SEMI = "DM Sans 14pt SemiBold"

FOOTER = "Mastering Agency Commercials  |  WTA  |  Maple Ranch"

# Template layouts used (index within the single WTA master)
L_TITLE_SWIRL = 0      # "Custom Layout": blue-red swirl gradient
L_DIVIDER = 5          # "1_Custom Layout": blue-red gradient with title + section name
L_CLOSING = 12         # "Title Slide": blue gradient, centred title, rule, body
L_DARK = 13            # "2_Title Slide": blue gradient, section name only
L_PLAIN = 4            # "5_Section Header": white, section name only
L_CONTENT = 32         # "23_Section Header": white, section name + title

prs = Presentation(str(TEMPLATE))
LAYOUTS = prs.slide_masters[0].slide_layouts

# --- remove the template's sample slides ---------------------------------------------
sldIdLst = prs.slides._sldIdLst
for sldId in list(sldIdLst):
    prs.part.drop_rel(sldId.rId)
    sldIdLst.remove(sldId)


# --- helpers -------------------------------------------------------------------------
def _inherit_footer(slide, layout):
    """python-pptx skips footer / slide-number placeholders; clone them from the layout."""
    for ph in layout.placeholders:
        t = ph.placeholder_format.type
        if t not in (PH.FOOTER, PH.SLIDE_NUMBER):
            continue
        el = copy.deepcopy(ph._element)
        slide.shapes._spTree.append(el)
        if t == PH.FOOTER:
            ext = el.find(qn("p:spPr")).find(qn("a:xfrm")).find(qn("a:ext"))
            ext.set("cx", str(int(Inches(8.0))))
            txBody = el.find(qn("p:txBody"))
            for p in txBody.findall(qn("a:p"))[1:]:
                txBody.remove(p)
            p = txBody.find(qn("a:p"))
            for r in p.findall(qn("a:r")) + p.findall(qn("a:fld")):
                p.remove(r)
            r = etree.SubElement(p, qn("a:r"))
            etree.SubElement(r, qn("a:rPr"), lang="en-US", dirty="0")
            etree.SubElement(r, qn("a:t")).text = FOOTER
            end = p.find(qn("a:endParaRPr"))
            if end is not None:
                p.remove(end)
                p.append(end)


def new_slide(layout_idx, eyebrow=None, title=None, notes=None, keep=()):
    layout = LAYOUTS[layout_idx]
    s = prs.slides.add_slide(layout)
    for ph in list(s.placeholders):
        idx = ph.placeholder_format.idx
        if ph.placeholder_format.type in (PH.TITLE, PH.CENTER_TITLE) and title is not None:
            ph.text_frame.text = title
            if layout_idx == L_CONTENT:
                # Fixed title frame: one line at 36pt, top-anchored, clear of the
                # section label above and of the sub-head/content zone below (y >= 1.75")
                ph.left, ph.top, ph.width, ph.height = Inches(0.52), Inches(0.88), Inches(12.29), Inches(0.72)
                tf = ph.text_frame
                tf.word_wrap = True
                tf.vertical_anchor = MSO_ANCHOR.TOP
                for r in tf.paragraphs[0].runs:
                    r.font.size = Pt(36)
        elif idx == 36 and eyebrow is not None:
            ph.text_frame.text = eyebrow
        elif idx in keep:
            continue
        else:
            ph._element.getparent().remove(ph._element)
    _inherit_footer(s, layout)
    if notes:
        s.notes_slide.notes_text_frame.text = notes.strip()
    return s


def ph(slide, idx):
    for p in slide.placeholders:
        if p.placeholder_format.idx == idx:
            return p
    raise KeyError(idx)


def _style_run(run, size, color, font, bold=False, italic=False):
    f = run.font
    f.size = Pt(size)
    f.name = font
    f.bold = bold
    f.italic = italic
    if color is not None:
        f.color.rgb = color


def text(slide, x, y, w, h, paras, size=14, color=TEXT, font=BODY, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, after=0, line=None, name=None):
    """paras: str | list of str | list of dict(t=..., size, color, font, bold, italic, align, after)
    A dict's t may itself be a list of run dicts for mixed styling."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    if isinstance(paras, str):
        paras = [paras]
    for i, p in enumerate(paras):
        if not isinstance(p, dict):
            p = {"t": p}
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = p.get("align", align)
        para.space_after = Pt(p.get("after", after))
        if p.get("before"):
            para.space_before = Pt(p["before"])
        if line or p.get("line"):
            para.line_spacing = p.get("line", line)
        runs = p["t"] if isinstance(p["t"], list) else [{"t": p["t"]}]
        for rd in runs:
            r = para.add_run()
            r.text = rd["t"]
            _style_run(r, rd.get("size", p.get("size", size)), rd.get("color", p.get("color", color)),
                       rd.get("font", p.get("font", font)), rd.get("bold", p.get("bold", bold)),
                       rd.get("italic", p.get("italic", False)))
        if p.get("bullet"):
            pPr = para._p.get_or_add_pPr()
            pPr.set("marL", str(int(Inches(0.22))))
            pPr.set("indent", str(-int(Inches(0.22))))
            for tag in ("a:buNone", "a:buChar", "a:buAutoNum"):
                for el in pPr.findall(qn(tag)):
                    pPr.remove(el)
            buClr = etree.SubElement(pPr, qn("a:buClr"))
            etree.SubElement(buClr, qn("a:srgbClr"), val=str(p.get("bucolor", RED)))
            etree.SubElement(pPr, qn("a:buFont"), typeface="Arial")
            etree.SubElement(pPr, qn("a:buChar"), char="•")
    return tb


def box(slide, x, y, w, h, fill=GREY, line=None, shape=MSO_SHAPE.RECTANGLE, name=None, line_w=1):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        s.name = name
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(line_w)
    s.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = 0.06
    s.text_frame.text = ""
    return s


def rule(slide, x1, y1, x2, y2, color=RULE, width=1, dash=False):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = color
    c.line.width = Pt(width)
    if dash:
        c.line.dash_style = 4  # dash
    return c


def label_in(shape, t, size=14, color=WHITE, font=SEMI, align=PP_ALIGN.CENTER, bold=False):
    tf = shape.text_frame
    tf.margin_left = tf.margin_right = Inches(0.1)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = t
    _style_run(r, size, color, font, bold)


def num_circle(slide, x, y, d, n, fill=NAVY, color=WHITE, size=16):
    c = box(slide, x, y, d, d, fill=fill, shape=MSO_SHAPE.OVAL)
    label_in(c, n, size=size, color=color, font=SEMI)
    c.text_frame.margin_left = c.text_frame.margin_right = 0
    return c


# Running clock for speaker-note timing cues
_clock = [0.0]


def cue(minutes):
    a = _clock[0]
    _clock[0] += minutes

    def fmt(m):
        return f"{int(m):02d}:{int(round((m - int(m)) * 60)):02d}"
    return f"[{fmt(a)} – {fmt(_clock[0])}]  ({minutes:g} min)\n\n"


# Layout grid for content slides (inches): section label 0.42, title 0.88–1.60,
# optional sub-head 1.75–2.10, content 2.35–6.30, footer rule 6.55.
SUB_Y = 1.75
TOP = 2.35


def subhead(s, t):
    text(s, 0.52, SUB_Y, 12.29, 0.35, t, size=16, color=MUTED)


def divider(part, title, sub, notes, minutes):
    s = new_slide(L_DIVIDER, eyebrow=part, title=title, notes=cue(minutes) + notes)
    t = ph(s, 0)
    t.left, t.top, t.width, t.height = Inches(0.47), Inches(2.45), Inches(11.5), Inches(1.3)
    t.text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
    t.text_frame.word_wrap = False
    for p in t.text_frame.paragraphs:
        for r in p.runs:
            _style_run(r, 54, WHITE, LIGHT)
    text(s, 0.49, 4.0, 10.0, 0.6, sub, size=22, color=WHITE, font=BODY)
    return s


# =====================================================================================
# PART 1 — FINANCE 101: HOW OUR AGENCY MAKES MONEY (00:00 – 15:00)
# =====================================================================================

# Cover -------------------------------------------------------------------------------
s = new_slide(L_TITLE_SWIRL, keep=(0,), notes=cue(1) + """
Welcome everyone, and thank Eric and the WTA leadership team for making time for this session.

Frame it: this is a practical, skill-building hour. The goal is to give account teams the commercial confidence to manage their accounts profitably, without guesswork.
""")
t = ph(s, 0)
t.left, t.top, t.width, t.height = Inches(0.52), Inches(1.7), Inches(11.0), Inches(1.2)
t.text_frame.word_wrap = False
t.text_frame.text = ""
r = t.text_frame.paragraphs[0].add_run()
r.text = "Mastering Agency Commercials"
_style_run(r, 54, WHITE, LIGHT)
text(s, 0.52, 3.05, 9.0, 0.9, "Aligning great client work with financial health", size=24, color=WHITE)
rule(s, 0.52, 4.25, 1.6, 4.25, color=WHITE, width=1)
text(s, 0.52, 4.45, 8.6, 0.8, [
    {"t": "Erin Jemas", "font": SEMI, "size": 16, "color": WHITE},
    {"t": "Founder & Principal, Price It Right", "size": 14, "color": WHITE},
], line=1.15)

# 1 — Welcome & objectives ------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="WELCOME", title="Today’s goals", notes=cue(2) + """
Three goals for the hour:
1. Understand how WTA makes money, and where profit comes from.
2. Learn the direct levers account leads control.
3. Get an introduction to WTA's evolved commercial model for FY2027.

Walk the agenda on the right briefly. Reassure the room: no accounting background needed. Everything builds from one simple equation.
""")
goals = ["Understand how WTA makes money, and where profit comes from",
         "Learn the direct levers account leads control",
         "Get an introduction to WTA’s evolved commercial model for FY2027"]
for i, g in enumerate(goals):
    y = TOP + i * 1.3
    box(s, 0.52, y, 7.2, 1.1, fill=GREY)
    num_circle(s, 0.82, y + 0.25, 0.6, f"0{i+1}", fill=NAVY, size=14)
    text(s, 1.7, y, 5.8, 1.1, g, size=18, color=NAVY, anchor=MSO_ANCHOR.MIDDLE, line=1.1)
box(s, 8.1, TOP, 4.71, 3.7, fill=NAVY)
text(s, 8.45, TOP + 0.2, 4.0, 0.35, "AGENDA  ·  60 MINUTES", size=12, color=TEAL, font=SEMI)
agenda = [("00:00", "Finance 101"), ("15:00", "WTA in Action"), ("35:00", "Commercial Model Intro"),
          ("55:00", "Q&A and wrap-up")]
for i, (when, what) in enumerate(agenda):
    y = TOP + 0.7 + i * 0.7
    text(s, 8.45, y, 1.0, 0.5, when, size=14, color=TEAL, font=SEMI, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 9.55, y, 3.1, 0.5, what, size=16, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

divider("PART 1  ·  00:00 – 15:00", "Finance 101", "How our agency makes money", """
Goal for this part: demystify agency accounting, agree on a common vocabulary, and define what financial health looks like.
(Terminology reviewed with Eric so the room hears the same language Finance uses.)
""", 0.5)

# 2 — The core equation ---------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="FINANCE 101", title="The core equation of business", notes=cue(3) + """
Every business, including ours, comes down to one equation: Revenue minus Expenses equals Profit.

Three terms to get right:
• Gross Revenue (billings): everything we invoice the client, including pass-through expenses.
• Pass-through costs: out-of-pocket costs passed directly to the client, such as media spend, production vendors, research and client travel. We bill them, but they were never ours to keep.
• Net Revenue (also called Net Sales or Gross Margin): gross billings minus pass-through costs. This is our true revenue, the money WTA actually earns to run the business.

A common mistake is to think everything we bill is revenue. It isn't. Every margin we discuss today is measured against Net Revenue.
""")
eq = [("Revenue", NAVY, WHITE), ("Expenses (costs)", GREY, NAVY), ("Profit", RED, WHITE)]
tw, ops = 3.45, 0.97
for i, (name, fill, fg) in enumerate(eq):
    x = 0.52 + i * (tw + ops)
    b = box(s, x, TOP, tw, 1.0, fill=fill)
    label_in(b, name, size=24, color=fg, font=BODY)
    if i < 2:
        text(s, x + tw, TOP, ops, 1.0, "−" if i == 0 else "=", size=40, color=NAVY, font=LIGHT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.52, 3.65, 12.29, 0.3, "KEY TERMS", size=12, color=RED, font=SEMI)
terms = [("Gross Revenue (billings)", "Everything we invoice to the client, including pass-through expenses"),
         ("Pass-through costs", "Out-of-pocket costs passed directly to the client: media, production vendors, research, client travel"),
         ("Net Revenue", "Billings minus pass-through. Our true revenue: the money WTA actually earns to run the business")]
cw, gap = 3.93, 0.25
for i, (head, body_) in enumerate(terms):
    x = 0.52 + i * (cw + gap)
    box(s, x, 4.05, cw, 2.2, fill=GREY if i < 2 else NAVY)
    fg = NAVY if i < 2 else WHITE
    text(s, x + 0.3, 4.25, cw - 0.6, 0.4, head, size=17, color=fg, font=SEMI)
    text(s, x + 0.3, 4.75, cw - 0.6, 1.4, body_, size=14, color=TEXT if i < 2 else WHITE, line=1.15)

# 3 — Direct costs vs overhead --------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="FINANCE 101", title="Expenses: direct costs vs. overhead", notes=cue(3) + """
Expenses come in two kinds.

DIRECT COSTS (client-related):
• Direct labor: salaries, benefits and payroll taxes for the hours our billable talent spends on client work. In practice: the share of a person's time on your client multiplied by their total compensation.
• Non-billable client out-of-pockets: client-related expenses we can't bill back, such as unapproved travel, entertainment and write-offs.

INDIRECT COSTS / OVERHEAD (agency operations): rent, utilities, IT and equipment, administrative staff, legal, HR and marketing.

KEY LINE: "Overhead is shared across the agency and managed by Eric and the finance team. What account leads manage on their day-to-day business is direct cost."
""")
box(s, 0.52, TOP, 6.0, 3.95, fill=NAVY)
text(s, 0.85, TOP + 0.2, 5.4, 0.35, "DIRECT COSTS  ·  CLIENT-RELATED", size=14, color=TEAL, font=SEMI)
text(s, 0.85, TOP + 0.75, 5.4, 3.0, [
    {"t": "Direct labor", "font": SEMI, "size": 18, "after": 3},
    {"t": "Salaries, benefits and payroll taxes for the hours our billable talent spends on client work", "after": 18},
    {"t": "Non-billable client out-of-pockets", "font": SEMI, "size": 18, "after": 3},
    {"t": "Client-related expenses we can’t bill back: unapproved travel, entertainment, write-offs"},
], size=14, color=WHITE, line=1.15)
box(s, 6.81, TOP, 6.0, 3.95, fill=GREY)
text(s, 7.14, TOP + 0.2, 5.4, 0.35, "INDIRECT COSTS / OVERHEAD  ·  AGENCY OPERATIONS", size=14, color=MUTED, font=SEMI)
text(s, 7.14, TOP + 0.75, 5.4, 3.0, [
    {"t": t_, "bullet": True, "after": 7, "bucolor": MUTED} for t_ in
    ["Rent and utilities", "IT and equipment", "Administrative staff", "Legal and HR", "Marketing"]
], size=16, color=TEXT)

# 4 — Benchmark metrics ---------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="FINANCE 101", title="The benchmarks of financial health", notes=cue(3) + """
Two targets every account lead should know.

Direct Margin (also called Contribution Margin) = Net Revenue − Direct Costs. Target: 55%–62% of Net Revenue.
Operating Profit = Direct Margin − Overhead. Target: about 20% of Net Revenue.

Why Direct Margin matters: if an account delivers $100K in Net Revenue, direct costs should land between $38K and $45K. That leaves $55K–$62K of Direct Margin to cover agency overhead and create profit.

Why profit matters: it pays for raises, bonuses, new technology, training and team perks.
""")
cards_ = [("DIRECT MARGIN", "Net Revenue  −  Direct Costs", "55–62%", "of Net Revenue", RED),
          ("OPERATING PROFIT", "Direct Margin  −  Overhead", "~20%", "of Net Revenue", NAVY)]
for i, (head, formula, tgt, unit, col) in enumerate(cards_):
    x = 0.52 + i * 4.25
    box(s, x, TOP, 4.0, 2.75, fill=GREY)
    text(s, x + 0.3, TOP + 0.2, 3.4, 0.35, head, size=14, color=col, font=SEMI)
    text(s, x + 0.3, TOP + 0.6, 3.4, 0.45, formula, size=16, color=NAVY)
    rule(s, x + 0.3, TOP + 1.2, x + 3.7, TOP + 1.2, color=RULE)
    text(s, x + 0.3, TOP + 1.3, 3.4, 0.25, "TARGET", size=11, color=MUTED, font=SEMI)
    text(s, x + 0.3, TOP + 1.55, 3.4, 0.85, tgt, size=48, color=col, font=LIGHT)
    text(s, x + 0.3, TOP + 2.35, 3.4, 0.3, unit, size=12, color=MUTED)
# $100K example
ex = 9.02
text(s, ex, TOP, 3.79, 0.35, "EXAMPLE: $100K OF NET REVENUE", size=12, color=MUTED, font=SEMI)
bar_y, bar_h = TOP + 0.5, 0.75
box(s, ex, bar_y, 3.79 * 0.42, bar_h, fill=NAVY)
box(s, ex + 3.79 * 0.42, bar_y, 3.79 * 0.58, bar_h, fill=RED)
text(s, ex, bar_y + 0.9, 1.6, 0.7, [{"t": "$38–45K", "font": SEMI, "color": NAVY, "size": 16},
                                     {"t": "Direct costs", "size": 13}], line=1.05)
text(s, ex + 1.75, bar_y + 0.9, 2.0, 0.7, [{"t": "$55–62K", "font": SEMI, "color": RED, "size": 16},
                                            {"t": "Direct Margin", "size": 13}], line=1.05)
text(s, ex, bar_y + 1.75, 3.79, 0.6, "Direct Margin covers agency overhead and creates profit",
     size=13, color=TEXT, line=1.1)
# why profit matters
text(s, 0.52, 5.35, 12.29, 0.3, "WHY PROFIT MATTERS  ·  IT PAYS FOR", size=12, color=MUTED, font=SEMI)
perks = ["Raises", "Bonuses", "New technology", "Training", "Team perks"]
pw = (12.29 - 4 * 0.2) / 5
for i, p_ in enumerate(perks):
    b = box(s, 0.52 + i * (pw + 0.2), 5.7, pw, 0.55, fill=NAVY)
    label_in(b, p_, size=14, color=WHITE, font=SEMI)

# 5 — Revenue vs profitability --------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="FINANCE 101", title="Quick example: revenue vs. profitability", notes=cue(2.5) + """
Ask the room first: "Which company would you rather run?" Take a few answers.

• Company Alpha: $2.0M revenue, $1.0M direct labor, $288K profit — a 14% margin.
• Company Beta: $250K revenue, $100K direct labor, $61K profit — a 24% margin.

KEY LINE: "Company Alpha makes more total dollars, but Company Beta is converting a healthier percentage to profit because their scope and staffing were managed properly. Profit margin is our true measure of health."
""")
cos = [("COMPANY ALPHA", "$2.0M", "$1.0M", "$288K", "14%", NAVY),
       ("COMPANY BETA", "$250K", "$100K", "$61K", "24%", RED)]
for i, (name, rev, dl, prof, mg, col) in enumerate(cos):
    x = 0.52 + i * 6.29
    box(s, x, TOP, 6.0, 0.5, fill=col)
    text(s, x + 0.3, TOP, 5.4, 0.5, name, size=14, color=WHITE, font=SEMI, anchor=MSO_ANCHOR.MIDDLE)
    box(s, x, TOP + 0.5, 6.0, 2.6, fill=GREY)
    for k, (lab, val) in enumerate([("Revenue", rev), ("Direct labor", dl), ("Profit", prof)]):
        text(s, x + 0.3 + k * 1.3, TOP + 0.7, 1.25, 0.3, lab, size=12, color=MUTED)
        text(s, x + 0.3 + k * 1.3, TOP + 0.98, 1.25, 0.45, val, size=20, color=NAVY, font=SEMI)
    text(s, x + 4.2, TOP + 0.7, 1.6, 0.3, "Profit margin", size=12, color=MUTED)
    text(s, x + 4.2, TOP + 0.9, 1.6, 0.9, mg, size=48, color=col, font=LIGHT)
    rule(s, x + 0.3, TOP + 1.95, x + 5.7, TOP + 1.95, color=RULE)
    text(s, x + 0.3, TOP + 2.1, 5.4, 0.8,
         "More total dollars, lower margin" if i == 0 else "Fewer dollars, healthier margin: scope and staffing managed well",
         size=14, color=TEXT, line=1.1)
box(s, 0.52, 5.45, 12.29, 0.8, fill=NAVY)
text(s, 0.85, 5.45, 11.7, 0.8, [
    {"t": [{"t": "Profit margin is our true measure of health,  ", "font": SEMI},
           {"t": "not total revenue.", "font": LIGHT}]}],
    size=20, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)


# =====================================================================================
# PART 2 — WTA IN ACTION: MANAGING GROWTH & MARGIN (15:00 – 35:00)
# =====================================================================================
divider("PART 2  ·  15:00 – 35:00", "WTA in Action", "Managing growth and margin", """
Goal for this part: connect the metrics to account leadership ownership, the levers account leads control, and dynamic FY2027 forecasting.
""", 0.5)

# 6 — Who owns financial health -------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="WTA IN ACTION", title="Who owns financial health at WTA?", notes=cue(3.5) + """
The account lead is the commercial owner of their client engagement.

Finance is your co-pilot and advisor, but account leads manage the daily scope, team mix and client relationships that create healthy margins.

Frame this as empowerment: financial management isn't back-office paperwork. It's how account leads protect their teams from overburn and make sure WTA is fairly paid for its work.
""")
box(s, 0.52, TOP, 12.29, 1.2, fill=NAVY)
text(s, 0.85, TOP, 11.7, 1.2, [
    {"t": [{"t": "The account lead ", "font": SEMI}, {"t": "is the commercial owner of their client engagement."}]}],
    size=24, color=WHITE, font=LIGHT, anchor=MSO_ANCHOR.MIDDLE)
own = [("ACCOUNT LEADS MANAGE", ["Day-to-day scope", "Team mix", "Client relationships"], RED),
       ("FINANCE PROVIDES", ["Co-pilot and advisor", "Reporting and analysis", "Pricing and forecasting support"], BLUE)]
for i, (head, items, col) in enumerate(own):
    x = 0.52 + i * 4.25
    box(s, x, 3.8, 4.0, 2.45, fill=GREY)
    text(s, x + 0.3, 4.0, 3.4, 0.35, head, size=14, color=col, font=SEMI)
    text(s, x + 0.3, 4.45, 3.4, 1.7, [{"t": t_, "bullet": True, "after": 6, "bucolor": col} for t_ in items],
         size=16, color=NAVY)
box(s, 9.02, 3.8, 3.79, 2.45, fill=None, line=NAVY, line_w=1.25)
text(s, 9.3, 3.8, 3.25, 2.45, [
    {"t": "WHY IT MATTERS", "size": 14, "font": SEMI, "color": RED, "after": 8},
    {"t": "Protect your team from overburn, and make sure WTA is fairly paid for its work.", "size": 16}],
    color=NAVY, anchor=MSO_ANCHOR.MIDDLE, line=1.15)

# 7 — Seven levers --------------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="WTA IN ACTION", title="Seven levers account leads control", notes=cue(6.5) + """
Seven practical levers to protect Direct Margin. Ask for a real example from the room on two or three of them.

1. Proper scoping and SOW detail: define deliverables, revision rounds and client assumptions clearly.
2. Never discount fee without scope: if the budget drops, reduce deliverables. Never work the same scope for less fee.
3. Staff mix alignment: use the right seniority for the task. Don't spend senior strategy hours on routine production.
4. Active scope-creep management: spot out-of-scope requests immediately, then issue a change order or swap deliverables.
5. Resource planning: work with project managers to track actuals against staff plans monthly.
6. Controlling non-billables: keep client travel and unapproved expenses to zero.
7. Contract hygiene: SOWs and POs signed before kickoff, to eliminate unbillable work-in-progress (WIP) and write-offs.
""")
subhead(s, "Practical ways to protect Direct Margin, every day")
levers = [
    ("Proper scoping & SOW detail", "Define deliverables, revision rounds and client assumptions clearly"),
    ("Never discount fee without scope", "If budget drops, reduce deliverables, never the fee alone"),
    ("Staff mix alignment", "Right seniority for the task; senior hours on senior work"),
    ("Active scope-creep management", "Flag out-of-scope requests; change order or swap deliverables"),
    ("Resource planning", "Track actuals against staff plans monthly with project managers"),
    ("Controlling non-billables", "Keep client travel and unapproved expenses to zero"),
    ("Contract hygiene", "SOWs and POs signed before kickoff: no unbillable WIP or write-offs"),
]
cw, rh = 6.0, 0.88
for i, (head, body_) in enumerate(levers):
    col_, row = (0, i) if i < 4 else (1, i - 4)
    x = 0.52 + col_ * 6.29
    y = TOP + row * (rh + 0.12)
    box(s, x, y, cw, rh, fill=GREY)
    num_circle(s, x + 0.22, y + 0.19, 0.5, str(i + 1), fill=RED if i % 2 == 0 else NAVY, size=14)
    text(s, x + 0.95, y + 0.1, cw - 1.15, rh - 0.2, [
        {"t": head, "font": SEMI, "color": NAVY, "size": 15, "after": 1},
        {"t": body_, "size": 13}], color=TEXT, anchor=MSO_ANCHOR.MIDDLE, line=1.05)
y = TOP + 3 * (rh + 0.12)
b = box(s, 6.81, y, cw, rh, fill=NAVY)
text(s, 7.11, y, cw - 0.6, rh, [
    {"t": [{"t": "The common thread  ", "font": SEMI, "color": TEAL},
           {"t": "clear scope, the right team, and signed paperwork"}]}],
    size=14, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, line=1.1)

# 8 — Forecasting ---------------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="WTA IN ACTION", title="Revenue forecasting: top-down vs. bottoms-up", notes=cue(5) + """
• Top-down forecasting: executive revenue targets set at the agency level.
• Bottoms-up forecasting: account leads building realistic, monthly revenue projections based on actual client SOWs, known pipeline and renewal dates.

KEY LINE: "Bottoms-up forecasting gives Eric and leadership a true picture of resource needs. It allows the agency to make smart hiring decisions, manage freelance costs and allocate talent effectively."

The two meet in the middle: leadership sets the targets, and account leads show how each client will actually get there, month by month.
""")
cols = [("TOP-DOWN", "Set by leadership", "Executive revenue targets set at the agency level", NAVY, WHITE, TEAL),
        ("BOTTOMS-UP", "Built by account leads", "Realistic monthly revenue projections based on actual client SOWs, known pipeline and renewal dates", GREY, NAVY, RED)]
for i, (head, who, body_, fill, fg, acc) in enumerate(cols):
    x = 0.52 + i * 6.29
    box(s, x, TOP, 6.0, 2.5, fill=fill)
    text(s, x + 0.35, TOP + 0.2, 5.3, 0.35, head, size=14, color=acc, font=SEMI)
    text(s, x + 0.35, TOP + 0.6, 5.3, 0.45, who, size=22, color=fg)
    text(s, x + 0.35, TOP + 1.2, 5.3, 1.2, body_, size=16, color=fg, line=1.15)
text(s, 0.52, 5.1, 12.29, 0.3, "WHAT A GOOD BOTTOMS-UP FORECAST MAKES POSSIBLE", size=12, color=MUTED, font=SEMI)
wins = ["Smart hiring decisions", "Managed freelance costs", "Effective talent allocation"]
ww = (12.29 - 2 * 0.25) / 3
for i, w_ in enumerate(wins):
    b = box(s, 0.52 + i * (ww + 0.25), 5.45, ww, 0.8, fill=None, line=NAVY, line_w=1.25)
    label_in(b, w_, size=16, color=NAVY, font=SEMI)

# 9 — Portfolio categories ------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="WTA IN ACTION", title="Managing FY2027 portfolio targets", notes=cue(4.5) + """
FY2027 targets are managed dynamically, account by account, and each account falls into a category with its own focus.

• Growth accounts: identify organic upsell opportunities and introduce new capabilities.
• Maintenance accounts: standardize delivery, protect baseline margins and streamline workflow.
• Declining / attrition accounts: actively manage resource roll-offs so labor costs don't outpace falling revenue.
• New business: set up the commercial terms correctly from day one.

Suggested follow-up: each account lead places every client in one of these four categories ahead of FY2027 planning.
""")
subhead(s, "Every account has a category, and every category has a focus")
quads = [
    ("Growth", "Identify organic upsell opportunities and introduce new capabilities", RED),
    ("Maintenance", "Standardize delivery, protect baseline margins and streamline workflow", NAVY),
    ("Declining / attrition", "Manage resource roll-offs so labor costs don’t outpace falling revenue", BLUE),
    ("New business", "Set up the commercial terms correctly from day one", ORANGE),
]
qw, qh = 6.0, 1.85
for i, (head, body_, col) in enumerate(quads):
    x = 0.52 + (i % 2) * (qw + 0.29)
    y = TOP + (i // 2) * (qh + 0.2)
    box(s, x, y, qw, qh, fill=GREY)
    num_circle(s, x + 0.35, y + 0.3, 0.6, f"0{i+1}", fill=col, size=14)
    text(s, x + 1.2, y + 0.3, qw - 1.5, 0.6, head, size=22, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + 1.2, y + 0.95, qw - 1.5, 0.8, body_, size=15, color=TEXT, line=1.15)


# =====================================================================================
# PART 3 — COMMERCIAL MODEL INTRODUCTION: EVOLVING BEYOND THE HOUR (35:00 – 55:00)
# =====================================================================================
divider("PART 3  ·  35:00 – 55:00", "Commercial Model Intro", "Evolving beyond the hour", """
Goal for this part: explain the industry shift away from time and materials toward standardized Solution Sets, with clear guardrails and a practical case study.
""", 0.5)

# 10 — Where the industry is going ----------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="COMMERCIAL MODEL", title="Where the industry is going", notes=cue(3) + """
• Current state: most of our work is time and materials (T&M) or retainers built from hours.
• The problem: T&M penalizes efficiency. When we get faster and smarter, we bill fewer hours. And clients scrutinize hourly rates instead of valuing the outcome.
• The evolved model: standardized, value-based Solution Sets that package defined deliverables with clear pricing.

The point to land: getting better at our craft should help the agency, not reduce what we're paid.
""")
steps3 = [("THE CURRENT STATE", "Time & materials and hourly-derived retainers", GREY, NAVY, MUTED),
          ("THE PROBLEM", "T&M penalizes efficiency: when we get faster and smarter, we bill fewer hours. Clients scrutinize rates instead of valuing the outcome.", GREY, NAVY, RED),
          ("THE EVOLVED MODEL", "Standardized, value-based Solution Sets: defined deliverables with clear pricing", NAVY, WHITE, TEAL)]
sw, sg = 3.75, 0.52
for i, (head, body_, fill, fg, acc) in enumerate(steps3):
    x = 0.52 + i * (sw + sg)
    box(s, x, TOP, sw, 3.9, fill=fill)
    text(s, x + 0.3, TOP + 0.25, sw - 0.6, 0.35, head, size=14, color=acc, font=SEMI)
    text(s, x + 0.3, TOP + 0.75, sw - 0.6, 3.0, body_, size=18 if i != 1 else 16, color=fg, line=1.15)
    if i < 2:
        box(s, x + sw + 0.08, TOP + 1.7, 0.36, 0.42, fill=RED, shape=MSO_SHAPE.RIGHT_ARROW)

# 11 — IS / IS NOT --------------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="COMMERCIAL MODEL", title="What this model is, and what it is not", notes=cue(3) + """
Be explicit here; this is where most of the questions come from.

• It IS: pre-packaged, modular scopes of work with fixed pricing, clear deliverables and standardized terms.
• It is NOT pay-for-performance: we are not tying agency compensation to client sales or contingency risk.
• It is NOT an à la carte discount menu: clients can't pick and choose individual micro-tasks at a discount. They purchase complete, integrated Solution Sets.
""")
box(s, 0.52, TOP, 6.0, 3.9, fill=NAVY)
num_circle(s, 0.85, TOP + 0.3, 0.6, "✓", fill=RED, size=18)
text(s, 1.65, TOP + 0.3, 4.6, 0.6, "It IS", size=24, color=WHITE, font=SEMI, anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.85, TOP + 1.2, 5.3, 2.5, [
    {"t": t_, "bullet": True, "after": 8, "bucolor": TEAL} for t_ in
    ["Pre-packaged, modular scopes of work", "Fixed pricing", "Clear deliverables", "Standardized terms"]],
    size=18, color=WHITE)
nots = [("Pay-for-performance", "We are not tying agency compensation to client sales or contingency risk"),
        ("An à la carte discount menu", "Clients purchase complete, integrated Solution Sets, not individual micro-tasks at a discount")]
for i, (head, body_) in enumerate(nots):
    y = TOP + i * 2.0
    box(s, 6.81, y, 6.0, 1.9, fill=GREY)
    num_circle(s, 7.11, y + 0.3, 0.5, "✕", fill=MUTED, size=14)
    text(s, 7.8, y + 0.22, 4.8, 0.65, [{"t": "It is NOT", "size": 13, "color": MUTED, "font": SEMI},
                                       {"t": head, "size": 18, "color": NAVY, "font": SEMI}], line=1.0)
    text(s, 7.8, y + 0.95, 4.8, 0.9, body_, size=14, color=TEXT, line=1.15)

# 12 — Four-step framework ------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="COMMERCIAL MODEL", title="How we turn services into Solution Sets", notes=cue(4) + """
A four-step framework.

1. Define the unit: what is the client actually buying? A social campaign, a brand platform, an activation, ongoing stewardship.
2. Define the scope: explicitly list what is included AND what is out of scope. Ambiguity is where margin disappears.
3. Price for value: use historical effort to benchmark costs, but set price based on the value of the deliverable.
4. Build the offering: present clean, menu-style proposals instead of complex staffing grids.
""")
subhead(s, "A four-step framework")
steps = [("Define the unit", "What is the client actually buying? A campaign, a platform, an activation, ongoing stewardship"),
         ("Define the scope", "List what is included and what is out of scope"),
         ("Price for value", "Historical effort benchmarks cost; deliverable value sets price"),
         ("Build the offering", "Clean, menu-style proposals instead of complex staffing grids")]
sw = 2.9
xs = [0.52 + i * (sw + 0.23) for i in range(4)]
rule(s, xs[0] + 0.4, TOP + 0.4, xs[-1] + 0.4, TOP + 0.4, color=RULE, width=1.5)
for i, (head, body_) in enumerate(steps):
    x = xs[i]
    num_circle(s, x, TOP, 0.8, f"0{i+1}", fill=NAVY if i < 3 else RED, size=18)
    text(s, x, TOP + 1.05, sw, 0.45, head, size=20, color=NAVY, font=SEMI)
    text(s, x, TOP + 1.55, sw - 0.15, 1.5, body_, size=15, color=TEXT, line=1.15)
box(s, 0.52, 5.6, 12.29, 0.68, fill=GREY)
text(s, 0.85, 5.6, 11.7, 0.68, [
    {"t": [{"t": "Ambiguity is where margin disappears.  ", "font": SEMI, "color": RED},
           {"t": "Every undefined element is a future scope disagreement."}]}],
    size=18, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)

# 13 — Case study ---------------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="CASE STUDY", title="Case study: a retail pitch win", notes=cue(4.5) + """
Retail pitch example (keep the client unnamed on screen).

The legacy approach: a multi-year RFP answered with complex FTE staffing options — 17 to 26 FTEs across 5 departments, $7.4M to $10.1M. The client got bogged down auditing hours and rates.

The Solution Set approach:
• Pulse Campaigns: targeted multi-platform activations at $75,000 per campaign. The client adjusts volume easily, with zero rate negotiations.
• Milestone Activations: full multi-channel brand moments at $500,000 per activation.
• Monthly Stewardship: ongoing strategic continuity and reporting at $25,000 per month.

KEY LINE: "By packaging our capabilities into clear Solution Sets, the client understood exactly what their budget delivered. It removed friction around staffing, accelerated speed to sale, and protected our margin."
""")
box(s, 0.52, TOP, 4.6, 3.9, fill=GREY)
text(s, 0.85, TOP + 0.25, 4.0, 0.35, "BEFORE  ·  LEGACY APPROACH", size=14, color=MUTED, font=SEMI)
text(s, 0.85, TOP + 0.75, 4.0, 0.8, "$7.4M–$10.1M", size=40, color=NAVY, font=LIGHT)
text(s, 0.85, TOP + 1.6, 4.0, 2.2, [
    {"t": "17–26 FTEs across 5 departments", "font": SEMI, "color": NAVY, "after": 10},
    {"t": "A multi-year RFP answered with complex staffing options. The client got bogged down auditing hours and rates."}],
    size=15, color=TEXT, line=1.15)
box(s, 5.32, TOP + 1.72, 0.36, 0.42, fill=RED, shape=MSO_SHAPE.RIGHT_ARROW)
text(s, 5.88, TOP, 6.93, 0.35, "AFTER  ·  SOLUTION SET APPROACH", size=14, color=RED, font=SEMI)
prods = [("Pulse Campaigns", "$75K", "per campaign", "Targeted multi-platform activations; client adjusts volume with zero rate negotiations"),
         ("Milestone Activations", "$500K", "per activation", "Full multi-channel brand moments"),
         ("Monthly Stewardship", "$25K", "per month", "Ongoing strategic continuity and reporting")]
for i, (name, price, unit, body_) in enumerate(prods):
    y = TOP + 0.45 + i * 1.17
    box(s, 5.88, y, 6.93, 1.05, fill=NAVY if i == 0 else (BLUE if i == 1 else RED))
    text(s, 6.15, y, 2.3, 1.05, [{"t": price, "size": 30, "font": LIGHT}, {"t": unit, "size": 12}],
         color=WHITE, anchor=MSO_ANCHOR.MIDDLE, line=0.95)
    text(s, 8.45, y, 4.2, 1.05, [{"t": name, "size": 16, "font": SEMI, "after": 1}, {"t": body_, "size": 13}],
         color=WHITE, anchor=MSO_ANCHOR.MIDDLE, line=1.05)

# 14 — Benefits -----------------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="COMMERCIAL MODEL", title="Benefits for clients and for WTA", notes=cue(2.5) + """
A win-win.

For our clients: total price transparency, predictable outcomes, faster turnaround, and flexibility without re-contracting.

For account teams: faster proposals; no more debates over timesheets and headcount; simpler change orders; and teams deliver proven services repeatedly without reinventing the wheel each time.
""")
cols = [("FOR OUR CLIENTS", ["Total price transparency", "Predictable outcomes", "Faster turnaround",
                             "Flexibility without re-contracting"], GREY, NAVY, RED),
        ("FOR ACCOUNT TEAMS", ["Faster proposals", "No debates over timesheets or headcount", "Simpler change orders",
                               "Proven services, delivered repeatedly"], NAVY, WHITE, TEAL)]
for i, (head, items, fill, fg, acc) in enumerate(cols):
    x = 0.52 + i * 6.29
    box(s, x, TOP, 6.0, 3.9, fill=fill)
    text(s, x + 0.4, TOP + 0.3, 5.2, 0.35, head, size=14, color=acc, font=SEMI)
    for k, it in enumerate(items):
        y = TOP + 0.9 + k * 0.72
        num_circle(s, x + 0.4, y + 0.08, 0.42, "✓", fill=acc, color=WHITE if fill == GREY else NAVY, size=12)
        text(s, x + 1.05, y, 4.8, 0.6, it, size=18, color=fg, anchor=MSO_ANCHOR.MIDDLE)

# 15 — Rollout ------------------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="COMMERCIAL MODEL", title="How WTA is rolling this out", notes=cue(2.5) + """
A practical, phased roadmap.

1. Starting with new business: introducing packaged Solution Sets in new pitches and inbound RFPs, where we have a clean slate.
2. Client-by-client strategic transition: working with account leads to assess existing retainer clients and transition them to Solution Sets at contract renewal or SOW milestone dates.
3. Tools and support: pre-built scope templates, pricing guidance, and cross-functional support from Finance and Operations.

Reassure the room: nothing changes overnight on existing accounts, and no one builds pricing alone.
""")
subhead(s, "A practical, phased roadmap")
phases = [("01", "Start with new business", "Packaged Solution Sets in new pitches and inbound RFPs, where we have a clean slate", NAVY),
          ("02", "Client-by-client transition", "Assess existing retainer clients together and transition them at contract renewal or SOW milestone dates", BLUE),
          ("03", "Tools & support", "Pre-built scope templates, pricing guidance, and cross-functional support from Finance and Operations", RED)]
pw, gap = 3.93, 0.25
for i, (n, head, body_, col) in enumerate(phases):
    x = 0.52 + i * (pw + gap)
    box(s, x, TOP, pw, 3.9, fill=GREY)
    num_circle(s, x + 0.3, TOP + 0.3, 0.7, n, fill=col, size=16)
    text(s, x + 0.3, TOP + 1.25, pw - 0.6, 0.8, head, size=20, color=NAVY, font=SEMI, line=1.05)
    text(s, x + 0.3, TOP + 2.25, pw - 0.6, 1.55, body_, size=15, color=TEXT, line=1.15)


# =====================================================================================
# WRAP-UP & Q&A (55:00 – 60:00)
# =====================================================================================

# 16 — Key takeaways ------------------------------------------------------------------
s = new_slide(L_DARK, eyebrow="KEY TAKEAWAYS", notes=cue(2) + """
Three things to take back to your accounts:
1. Direct Margin (55%–62%) is our primary measure of account health.
2. Account leads control the levers: scope clarity, staff mix and change management.
3. Solution Sets help us sell clarity and value instead of just selling hours.
""")
rules_ = [("Direct Margin of 55–62%", "is our primary measure of account health."),
          ("Account leads control the levers:", "scope clarity, staff mix and change management."),
          ("Solution Sets help us sell clarity and value,", "instead of just selling hours.")]
for i, (a_, b_) in enumerate(rules_):
    y = 1.35 + i * 1.65
    text(s, 0.52, y, 1.4, 1.2, f"0{i+1}", size=60, color=WHITE, font=LIGHT, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 2.1, y, 10.5, 1.2, [
        {"t": a_, "font": SEMI, "size": 28, "color": WHITE, "after": 2},
        {"t": b_, "font": LIGHT, "size": 24, "color": WHITE}], anchor=MSO_ANCHOR.MIDDLE, line=1.0)
    if i < 2:
        rule(s, 2.1, y + 1.42, 12.81, y + 1.42, color=WHITE, width=0.5)

# 17 — Q&A ----------------------------------------------------------------------------
s = new_slide(L_CLOSING, title="Questions & open floor", keep=(16,), notes=cue(3) + """
Open the floor for account leads to ask practical questions about forecasting, SOW scoping or upcoming renewals. Eric (CFO) and commercial leads are available to add color.

Questions to be ready for:
• "Do we still fill out timesheets?" Yes. Internally they measure cost and margin; they just stop being how we price for clients.
• "What happens to my existing retainer clients?" Nothing changes overnight. We'll plan each transition together around renewal or SOW milestone dates.
• "Who sets the price of a Solution Set?" Account leadership and Finance together, using pricing guidance and scope templates.
""")
t = ph(s, 0)
t.left, t.top, t.width, t.height = Inches(0.52), Inches(2.2), Inches(12.29), Inches(1.2)
for r in t.text_frame.paragraphs[0].runs:
    r.font.size = Pt(54)
b = ph(s, 16)
b.text_frame.text = ""
p = b.text_frame.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "Forecasting  ·  SOW scoping  ·  Upcoming renewals"
_style_run(r, 18, WHITE, BODY)
p = b.text_frame.add_paragraph()
p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "With Eric (CFO) and commercial leads"
_style_run(r, 16, WHITE, LIGHT)

# Appendix — cheat sheet of terms -----------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="APPENDIX", title="Cheat sheet of terms", notes="""
Reference slide for after the session (not presented). Terms and targets aligned with Eric (CFO).
""")
rows = [("Metric / term", "Plain-language definition", "WTA target"),
        ("Net Revenue (Net Sales)", "Total billings minus pass-through expenses (media, external production, research)", "True agency top line"),
        ("Direct Labor", "Total compensation (salary, taxes, benefits) for billable time spent on an account", "Max 40% of Net Revenue"),
        ("Direct Margin", "What remains from Net Revenue after subtracting Direct Costs (labor + non-billable OOPs)", "55% – 62%"),
        ("Overhead", "Fixed agency operating expenses (rent, admin, IT, executive, shared services)", "~30% of Net Revenue"),
        ("Operating Profit", "What’s left after Direct Costs and Overhead are covered", "~20%"),
        ("Solution Set", "A standardized, packaged scope of work with defined inclusions, exclusions and fixed pricing", "Core WTA commercial model")]
tbl = s.shapes.add_table(len(rows), 3, Inches(0.52), Inches(TOP - 0.25), Inches(12.29), Inches(3.9)).table
tbl.columns[0].width = Inches(2.9)
tbl.columns[1].width = Inches(6.69)
tbl.columns[2].width = Inches(2.7)
for ri, row_ in enumerate(rows):
    tbl.rows[ri].height = Inches(0.55)
    for ci, val in enumerate(row_):
        c = tbl.cell(ri, ci)
        c.fill.solid()
        c.fill.fore_color.rgb = NAVY if ri == 0 else (GREY if ri % 2 else WHITE)
        c.margin_left = c.margin_right = Inches(0.15)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = c.text_frame
        tf.text = ""
        r = tf.paragraphs[0].add_run()
        r.text = val
        _style_run(r, 13 if ri else 13, WHITE if ri == 0 else (NAVY if ci != 1 else TEXT),
                   SEMI if (ri == 0 or ci == 0) else BODY)

# Sections for easy navigation in PowerPoint ---------------------------------------
SECTIONS = [("Part 1 – Finance 101", 0), ("Part 2 – WTA in Action", 7), ("Part 3 – Commercial Model Intro", 12),
            ("Wrap-up & Q&A", 19), ("Appendix", 21)]

P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
ids = [sid.get("id") for sid in prs.slides._sldIdLst]
pres_el = prs.part._element
extLst = pres_el.find(qn("p:extLst"))
if extLst is None:
    extLst = etree.SubElement(pres_el, qn("p:extLst"))
for ext in extLst.findall(qn("p:ext")):
    if ext.get("uri") == "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}":
        extLst.remove(ext)
ext = etree.Element(qn("p:ext"), uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}")
extLst.insert(0, ext)
secLst = etree.SubElement(ext, f"{{{P14}}}sectionLst", nsmap={"p14": P14})
import uuid
bounds = [i for _, i in SECTIONS] + [len(ids)]
for k, (name, start) in enumerate(SECTIONS):
    sec = etree.SubElement(secLst, f"{{{P14}}}section", name=name, id="{" + str(uuid.uuid4()).upper() + "}")
    lst = etree.SubElement(sec, f"{{{P14}}}sldIdLst")
    for sid in ids[start:bounds[k + 1]]:
        etree.SubElement(lst, f"{{{P14}}}sldId", id=sid)

prs.core_properties.title = "Mastering Agency Commercials"
prs.core_properties.author = "Erin Jemas, Price It Right"
prs.core_properties.subject = "Maple Ranch account leadership session"
prs.save(str(OUT))
print(f"Saved {OUT}  ({len(ids)} slides, timed to {_clock[0]:g} min)")
