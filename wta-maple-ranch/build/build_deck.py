"""Build the WTA Maple Ranch deck "Stop Selling Hours" on the WTA 2026 master template.

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
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE.parent / "WTA-Stop-Selling-Hours-Maple-Ranch.pptx"

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

FOOTER = "WTA Commercial Operating System  |  Leadership Retreat  |  Maple Ranch"

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
                ph.width = Inches(12.29)
                for r in ph.text_frame.paragraphs[0].runs:
                    r.font.size = Pt(40)
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


# =====================================================================================
# THE PROVOCATION (00:00 – 05:00)
# =====================================================================================

# 1 — Title ---------------------------------------------------------------------------
s = new_slide(L_TITLE_SWIRL, title=None, keep=(0,), notes=cue(2) + """
OPEN — set the stage.
• Thank Eric and the WTA leadership team for commissioning this work. This session is the start of how WTA will create, capture and scale value in FY2027.
• Introduce yourself briefly: 20+ years in agency finance and commercial leadership across holding-company, independent and PE-backed agencies.\n• Frame the why: this is about how WTA creates, captures and scales value, and an elevation in how the agency values its own expertise, internally and externally.
• Audience: WTA executives and select agency leadership. This session consolidates the finance, portfolio and commercial-model foundations into one hour.

THE HOOK (deliver slowly):
"You nailed the pitch. The creative was brilliant. And then procurement spent three weeks negotiating 25% of an Associate Creative Director.
That's not a creative problem. That's a commercial model problem."

Pause. Then: "For the next 60 minutes we're going to fix the commercial model."
""")
t = ph(s, 0)
t.left, t.top, t.width, t.height = Inches(0.52), Inches(1.55), Inches(9.5), Inches(2.4)
tf = t.text_frame
tf.text = ""
p = tf.paragraphs[0]
r = p.add_run(); r.text = "Stop selling "
_style_run(r, 66, WHITE, LIGHT)
r = p.add_run(); r.text = "hours"
_style_run(r, 66, WHITE, SEMI, italic=True)
text(s, 0.52, 3.85, 8.6, 0.9,
     "The WTA Commercial Operating System: how to productize agency services and price for value",
     size=20, color=WHITE, font=BODY, line=1.1)
rule(s, 0.52, 5.05, 1.6, 5.05, color=WHITE, width=1)
text(s, 0.52, 5.25, 8.6, 0.8, [
    {"t": "Erin Jemas", "font": SEMI, "size": 16, "color": WHITE},
    {"t": "Founder & Principal, Price It Right", "size": 14, "color": WHITE},
], line=1.15)

# 2 — Agenda --------------------------------------------------------------------------
s = new_slide(L_PLAIN, eyebrow="AGENDA  ·  60 MINUTES", notes=cue(0.5) + """
Quick roadmap — don't linger.
"We'll start with why the market is breaking the hourly model, build a shared financial vocabulary with Eric's team, show you what that means for how you run your portfolio, and then introduce the new commercial model and how we roll it out."
""")
rows = [
    ("01", "The Provocation", "Why pricing inputs is breaking", "00:00 – 05:00"),
    ("02", "Agency Economics 101", "The direct margin engine", "05:00 – 20:00"),
    ("03", "WTA in Action", "Front-line portfolio stewardship", "20:00 – 38:00"),
    ("04", "The Commercial Operating System", "Stop selling hours", "38:00 – 55:00"),
    ("05", "Roadmap & Q&A", "Our FY2027 commercial mandate", "55:00 – 60:00"),
]
y0, rh = 1.05, 1.02
for i, (n, name, sub, when) in enumerate(rows):
    y = y0 + i * rh
    text(s, 0.52, y, 1.0, 0.75, n, size=36, color=RED, font=LIGHT, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 1.55, y, 6.8, 0.75, [
        {"t": name, "size": 26, "color": NAVY, "font": BODY},
    ], anchor=MSO_ANCHOR.MIDDLE)
    text(s, 7.6, y, 3.4, 0.75, sub, size=14, color=MUTED, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 10.9, y, 1.9, 0.75, when, size=14, color=NAVY, font=SEMI, align=PP_ALIGN.RIGHT,
         anchor=MSO_ANCHOR.MIDDLE)
    rule(s, 0.52, y + 0.86, 12.81, y + 0.86, color=RULE)

# 3 — Market forces -------------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="THE PROVOCATION", title="Why time-based pricing is collapsing", notes=cue(2.5) + """
Three pressures are colliding at the same time — none of them are going away.

1. Economic scrutiny — every dollar is questioned. Clients want transparency and predictable ROI from agency partners.
2. Procurement influence — rate cards get treated as commodities to grind down. Time is no longer a defensible pricing unit when procurement compares us line by line with every other vendor.
3. AI & technology disruption — the time it takes to deliver brilliance is compressing.

KEY LINE: "When AI and modern expertise allow your team to do in two hours what used to take twenty, selling by the hour means efficiency is financial suicide."

Transition: "Today, we shift from selling labor inputs to pricing packaged impact. But first, we need a shared language for how this agency actually makes money."
""")
cards = [
    ("01", "Economic scrutiny", "Clients demand total transparency and predictable ROI on every dollar."),
    ("02", "Procurement influence", "Rate cards are treated as commodities to compare and grind down."),
    ("03", "AI & technology disruption", "The time it takes to deliver brilliance is compressing fast."),
]
cw, gap = 3.93, 0.25
for i, (n, head, body) in enumerate(cards):
    x = 0.52 + i * (cw + gap)
    box(s, x, 2.3, cw, 2.35, fill=GREY)
    text(s, x + 0.3, 2.5, 1.2, 0.6, n, size=32, color=RED, font=LIGHT)
    text(s, x + 0.3, 3.15, cw - 0.6, 0.4, head, size=18, color=NAVY, font=SEMI)
    text(s, x + 0.3, 3.62, cw - 0.6, 0.95, body, size=14, color=TEXT, line=1.15)
q = box(s, 0.52, 4.95, 12.29, 1.3, fill=NAVY)
text(s, 0.95, 5.05, 11.4, 1.1, [
    {"t": [{"t": "When your team can do in two hours what used to take twenty, "},
           {"t": "selling by the hour makes efficiency financial suicide.", "font": SEMI, "color": WHITE}]},
], size=20, color=WHITE, font=LIGHT, anchor=MSO_ANCHOR.MIDDLE, line=1.1)


# =====================================================================================
# PART 1 — AGENCY ECONOMICS 101 (05:00 – 20:00)
# =====================================================================================
def divider(part, title, sub, notes, minutes):
    s = new_slide(L_DIVIDER, eyebrow=part, title=title, notes=cue(minutes) + notes)
    t = ph(s, 0)
    t.left, t.top, t.width, t.height = Inches(0.47), Inches(2.45), Inches(11.5), Inches(1.5)
    t.text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
    t.text_frame.word_wrap = False
    for p in t.text_frame.paragraphs:
        for r in p.runs:
            _style_run(r, 54, WHITE, LIGHT)
    text(s, 0.49, 4.15, 8.0, 0.6, sub, size=22, color=WHITE, font=BODY)
    return s


divider("PART 1  ·  05:00 – 20:00", "Agency Economics 101", "The direct margin engine", """
Frame Part 1: "This is the universal financial vocabulary. It's aligned with Eric's finance team, so when you talk to Finance — and when Finance talks to you — we mean exactly the same thing by every term."
(Before Maple Ranch: Eric to review terminology on the next four slides for consistency with WTA reporting.)
""", 0.5)

# 4 — Anatomy of agency revenue -----------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="AGENCY ECONOMICS 101", title="Not every dollar we bill is agency revenue",
              notes=cue(3) + """
Clarify terminology — this is the most common misunderstanding in agency finance.

• Gross Billings: everything we invoice the client.
• Pass-through / out-of-pocket: media spend, third-party production, research, travel. We bill it and the client reimburses it — but it was never ours. It flows through.
• Net Revenue (Net Sales): what remains after pass-through. This is our true revenue baseline — every margin metric we discuss today is a percentage of Net Revenue, never of billings.

Illustrative example: bill a client $1.0M, of which $400K is media and production pass-through → our Net Revenue is $600K.

Ask the room: "On your biggest account, do you know the split between billings and Net Revenue?"
""")
tiles = [
    ("Gross Billings", "$1.0M", "Everything we invoice the client", NAVY),
    ("Pass-Through", "$0.4M", "Media, third-party production, research, travel — reimbursed, not earned", MUTED),
    ("Net Revenue", "$0.6M", "What remains. Our true revenue baseline — every margin starts here", RED),
]
tw, ops = 3.45, 0.97
for i, (name, val, desc, col) in enumerate(tiles):
    x = 0.52 + i * (tw + ops)
    b = box(s, x, 2.35, tw, 1.75, fill=col if i != 1 else GREY)
    fg = WHITE if i != 1 else NAVY
    text(s, x + 0.3, 2.5, tw - 0.6, 0.4, name.upper(), size=14, color=fg, font=SEMI)
    text(s, x + 0.3, 2.95, tw - 0.6, 0.95, val, size=48, color=fg, font=LIGHT)
    text(s, x, 4.3, tw, 0.9, desc, size=14, color=TEXT, line=1.15)
    if i < 2:
        text(s, x + tw, 2.35, ops, 1.75, "−" if i == 0 else "=", size=48, color=NAVY, font=LIGHT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.52, 5.6, 12.29, 0.6, [
    {"t": [{"t": "Rule of thumb  ", "font": SEMI, "color": NAVY},
           {"t": "Every margin metric in this session is measured against Net Revenue, never billings."}]},
], size=16, color=TEXT)
text(s, 0.52, 6.15, 6, 0.3, "Figures illustrative", size=10, color=MUTED)

# 5 — Direct costs vs overhead --------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="AGENCY ECONOMICS 101", title="What account leads control", notes=cue(3.5) + """
Two kinds of cost — and only one of them is yours.

DIRECT COSTS (in your control):
• Direct labor — the salary, payroll taxes and benefits of the billable talent working on your accounts. Practically: % of a person's time on your client × their total compensation. Staffing decisions ARE cost decisions.
• Non-billable client OOPs — whatever the client contract says we can't bill: local travel, entertainment, presentations, overtime meals/cars, unapproved travel — and write-offs (including reserves on aged WIP and A/R).

INDIRECT COSTS / OVERHEAD (Finance's control): rent, IT infrastructure, admin, HR, legal, executive, new-business/pitch costs.

KEY LINE: "Eric and the finance team manage the overhead. Your mandate as account leaders is managing the direct margin of your client engagements."
""")
# left: in your control (navy)
box(s, 0.52, 2.3, 6.0, 3.95, fill=NAVY)
text(s, 0.85, 2.5, 5.4, 0.35, "DIRECT COSTS  ·  IN YOUR CONTROL", size=14, color=TEAL, font=SEMI)
text(s, 0.85, 3.0, 5.4, 3.1, [
    {"t": "Direct labor", "font": SEMI, "size": 18, "after": 2},
    {"t": "Salaries, payroll taxes and benefits of the billable talent on your accounts", "after": 14},
    {"t": "Non-billable client out-of-pockets", "font": SEMI, "size": 18, "after": 2},
    {"t": "Unapproved or local travel, entertainment, presentations, overtime meals and cars", "after": 14},
    {"t": "Write-offs", "font": SEMI, "size": 18, "after": 2},
    {"t": "Unbilled WIP and aged receivables that never get collected"},
], size=14, color=WHITE, line=1.1)
# right: finance's control (grey)
box(s, 6.81, 2.3, 6.0, 3.95, fill=GREY)
text(s, 7.14, 2.5, 5.4, 0.35, "INDIRECT COSTS / OVERHEAD  ·  FINANCE’S CONTROL", size=14, color=MUTED, font=SEMI)
text(s, 7.14, 3.0, 5.4, 3.1, [
    {"t": t, "bullet": True, "after": 8, "bucolor": MUTED} for t in
    ["Rent and facilities", "IT infrastructure and equipment", "Administration and HR",
     "Legal and accounting", "Executive leadership", "New business and pitch costs"]
], size=16, color=TEXT)

# 6 — North star: DCM -----------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="AGENCY ECONOMICS 101", title="The north star: Direct Contribution Margin",
              notes=cue(4) + """
This is the one metric that tells you whether an engagement is healthy.

Direct Margin = Net Revenue − Direct Costs (direct labor + non-billable client OOPs). Express it as a % of Net Revenue.

WTA benchmark: 55%–62% Direct Margin. Operating profit goal: 20%–25% net.

Make the math simple: "For every $100K of Net Revenue, your direct labor and non-billable expenses should not exceed $38K–$45K."
Walk the bar on the right: at the 55% floor, $45K of direct cost leaves $55K of direct margin. Overhead takes roughly $35K of that, leaving the $20K operating profit that hits our goal. Slip below the floor and profit is the first thing to go — overhead doesn't shrink because one account ran hot.

KEY LINE: "When account leads consistently protect a 55%+ Direct Margin, the agency hits its profit goals — and that's what funds bonuses, raises, hiring and creative investment."

(Eric: confirm the ~35% overhead illustration matches WTA's actual overhead ratio.)
""")
text(s, 0.52, 2.2, 7.2, 0.35, "THE FORMULA", size=14, color=RED, font=SEMI)
text(s, 0.52, 2.6, 7.4, 0.7, [
    {"t": [{"t": "Direct Margin  =  ", "font": SEMI, "color": NAVY},
           {"t": "Net Revenue  −  Direct Costs", "color": NAVY}]},
], size=24, font=BODY)
rule(s, 0.52, 3.55, 7.3, 3.55, color=RULE)
stats = [("55–62%", "Direct Margin target", "of Net Revenue", RED),
         ("20–25%", "Operating profit goal", "net, after overhead", NAVY)]
for i, (v, l1, l2, col) in enumerate(stats):
    x = 0.52 + i * 3.6
    text(s, x, 3.8, 3.5, 1.0, v, size=54, color=col, font=LIGHT)
    text(s, x, 4.85, 3.4, 0.7, [{"t": l1, "font": SEMI, "color": NAVY, "size": 16}, {"t": l2, "color": MUTED}],
         size=14)
# $100K stacked bar (at the 55% floor)
bx, by, bw, bh = 9.25, 2.55, 1.25, 3.7
segs = [(45, NAVY, "Direct costs", "≤ $45K"), (35, RULE, "Overhead", "≈ $35K"), (20, RED, "Operating profit", "≥ $20K")]
yy = by
for v, col, name, amt in segs:
    h = bh * v / 100
    box(s, bx, yy, bw, h, fill=col)
    text(s, bx + bw + 0.2, yy + h / 2 - 0.3, 2.2, 0.6, [
        {"t": amt, "font": SEMI, "color": NAVY if col != RED else RED, "size": 16},
        {"t": name, "color": TEXT, "size": 14}], line=1.0)
    yy += h
text(s, bx - 0.6, by - 0.4, bw + 1.2, 0.3, "Every $100K of Net Revenue", size=12, color=MUTED,
     align=PP_ALIGN.CENTER)
# bracket for direct margin (overhead + profit)
top = by + bh * 0.45
bot = by + bh
rule(s, bx - 0.2, top, bx - 0.2, bot, color=RED, width=1.5)
rule(s, bx - 0.2, top, bx - 0.05, top, color=RED, width=1.5)
rule(s, bx - 0.2, bot, bx - 0.05, bot, color=RED, width=1.5)
text(s, bx - 1.95, (top + bot) / 2 - 0.3, 1.65, 0.6, [
    {"t": "≥ $55K", "font": SEMI, "color": RED, "size": 16},
    {"t": "Direct margin", "size": 14}], align=PP_ALIGN.RIGHT, line=1.0)

# 7 — Revenue vs profitability -------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="AGENCY ECONOMICS 101  ·  INTERACTIVE", title="Revenue is not the same as profitability",
              notes=cue(4) + """
Make it interactive: show the three profiles and ask the room, "Which account would you rather run?" Let a few people answer before revealing the margins.

• Account Alpha: $2.0M Net Revenue, $1.0M direct labor → 50% Direct Margin — below the 55% floor → 14% operating margin.
• Account Gamma: $1.0M Net Revenue, $450K direct labor → 55% Direct Margin — right at the floor → 19% operating margin.
• Account Beta: $250K Net Revenue, $100K direct labor → 60% Direct Margin — inside the band → 24% operating margin.
(Overhead is allocated at the same ~35.6% of Net Revenue to each account in this example.)

KEY LINE: "Account Alpha brings in 8x more top line than Beta, but Beta converts nearly double the profit percentage. In WTA's new operating model, we reward margin health and commercial discipline — not just top-line volume."

Then the agency-total line: one large, under-margined account drags the whole agency to 17% — below our 20% goal — even though two of three accounts are healthy.
""")
accts = [
    ("ACCOUNT ALPHA", "$2.0M", "$1.0M", "50%", "14%", False),
    ("ACCOUNT GAMMA", "$1.0M", "$450K", "55%", "19%", True),
    ("ACCOUNT BETA", "$250K", "$100K", "60%", "24%", True),
]
cw, gap = 3.93, 0.25
for i, (name, nr, dl, dm, om, ok) in enumerate(accts):
    x = 0.52 + i * (cw + gap)
    box(s, x, 2.2, cw, 0.5, fill=NAVY)
    text(s, x + 0.3, 2.2, cw - 0.6, 0.5, name, size=14, color=WHITE, font=SEMI, anchor=MSO_ANCHOR.MIDDLE)
    box(s, x, 2.7, cw, 2.75, fill=GREY)
    text(s, x + 0.3, 2.85, 1.7, 0.3, "Net Revenue", size=12, color=MUTED)
    text(s, x + 0.3, 3.1, 1.7, 0.45, nr, size=22, color=NAVY, font=SEMI)
    text(s, x + 2.0, 2.85, 1.7, 0.3, "Direct labor", size=12, color=MUTED)
    text(s, x + 2.0, 3.1, 1.7, 0.45, dl, size=22, color=NAVY, font=BODY)
    rule(s, x + 0.3, 3.72, x + cw - 0.3, 3.72, color=RULE)
    text(s, x + 0.3, 3.85, 1.7, 0.3, "Direct Margin", size=12, color=MUTED)
    text(s, x + 0.3, 4.1, 1.7, 0.8, dm, size=44, color=RED if not ok else NAVY, font=LIGHT)
    text(s, x + 2.0, 3.85, 1.7, 0.3, "Operating margin", size=12, color=MUTED)
    text(s, x + 2.0, 4.1, 1.7, 0.8, om, size=44, color=RED if not ok else NAVY, font=LIGHT)
    text(s, x + 0.3, 4.95, cw - 0.6, 0.35, "Below the 55% floor" if not ok else "Within target", size=12,
         color=RED if not ok else BLUE, font=SEMI)
text(s, 0.52, 5.7, 12.29, 0.6, [
    {"t": [{"t": "Agency total  ", "font": SEMI, "color": NAVY},
           {"t": "$3.25M Net Revenue  ·  52% Direct Margin  ·  17% operating margin. "},
           {"t": "One account pulls the agency below its 20% goal.", "font": SEMI, "color": RED}]},
], size=14, color=TEXT)
text(s, 0.52, 6.12, 9, 0.3, "Illustrative accounts; overhead allocated at the same share of Net Revenue to each",
     size=10, color=MUTED)


# =====================================================================================
# PART 2 — WTA IN ACTION (20:00 – 38:00)
# =====================================================================================
divider("PART 2  ·  20:00 – 38:00", "WTA in Action", "Front-line portfolio stewardship", """
"Now that we share the vocabulary, let's talk about who owns these numbers. The answer is: you do — with Finance and Operations right alongside you."
""", 0.5)

# 8 — Mindset shift -------------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="WTA IN ACTION", title="You are portfolio managers", notes=cue(3.5) + """
The shift in mindset for account leads — and for Finance.

• Account leads move from being passive recipients of quarterly finance reports to being active general managers of their client portfolio's margin. Revenue forecast and profitability targets are the gateway to evaluating the financial health of every client and portfolio — and you own them.
• Finance changes too: no longer an auditing body checking timesheets after the fact, but your strategic co-pilot — helping you price, scope and structure growth before it's sold.

There is cross-functional support behind every metric: Finance, Operations / resource management and project management partner with you on each account. Encourage every lead to meet their finance partner regularly — at minimum monthly.
""")
pairs = [
    ("ACCOUNT LEADS", "Passive recipient of quarterly finance reports",
     "Active general manager of your client portfolio’s margin"),
    ("FINANCE", "Auditor checking timesheets after the fact",
     "Strategic co-pilot helping you price and structure growth"),
]
text(s, 2.75, 2.15, 4.0, 0.35, "FROM", size=14, color=MUTED, font=SEMI)
text(s, 8.0, 2.15, 4.0, 0.35, "TO", size=14, color=RED, font=SEMI)
for i, (who, frm, to) in enumerate(pairs):
    y = 2.6 + i * 1.85
    text(s, 0.52, y, 2.0, 1.55, who, size=14, color=NAVY, font=SEMI, anchor=MSO_ANCHOR.MIDDLE)
    box(s, 2.75, y, 4.5, 1.55, fill=GREY)
    text(s, 3.05, y, 3.9, 1.55, frm, size=18, color=MUTED, anchor=MSO_ANCHOR.MIDDLE, line=1.1)
    a = box(s, 7.35, y + 0.55, 0.45, 0.45, fill=RED, shape=MSO_SHAPE.RIGHT_ARROW)
    box(s, 8.0, y, 4.81, 1.55, fill=NAVY)
    text(s, 8.3, y, 4.2, 1.55, to, size=18, color=WHITE, font=SEMI, anchor=MSO_ANCHOR.MIDDLE, line=1.1)

# 9 — Five levers ---------------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="WTA IN ACTION", title="Five levers of client profitability you control",
              notes=cue(6) + """
These are the levers that move Direct Margin. Walk each one with a real example from the room if you can.

1. Pricing with buffer — never discount without pulling scope. If the budget is lower, offer options: reduce deliverables, explain the value at the full price, or show a higher-value option worth finding funding for. Fixed fees carry a baseline buffer of at least 10%.
2. Staff mix & seniority alignment — match talent level to task complexity. Senior people on commodity production is the fastest way to burn margin. Monitor staff mix against the agreed staff plan.
3. Aggressive scope-creep management — know what's in and out of scope; flag out-of-scope requests the moment they happen; swap deliverables (keep the fee flat) or add deliverables and increase the fee. Never absorb unbilled hours. Document everything.
4. Contract & terms hygiene — signed estimate/SOW and PO before kickoff, every time (production included). This is how we eliminate WIP write-offs and aged receivables.
5. Dynamic utilization planning — partner with resource management to keep the pipeline visible, so people move to other work during client delays and project gaps instead of sitting under-utilized.
""")
levers = [
    ("Pricing with buffer", "Never discount without pulling scope. Fixed fees carry a baseline buffer of 10%+.", NAVY),
    ("Staff mix & seniority", "Match talent level to task complexity. No senior burn on commodity production.", BLUE),
    ("Scope-creep control", "Flag out-of-scope requests immediately. Swap deliverables, never work unbilled hours.", RED),
    ("Contract & terms hygiene", "Signed estimates and POs before kickoff. Eliminate WIP write-offs.", ORANGE),
    ("Utilization planning", "Partner with resource management so project gaps don’t become idle time.", CYAN),
]
cw, gap = 2.27, 0.235
for i, (head, body, col) in enumerate(levers):
    x = 0.52 + i * (cw + gap)
    box(s, x, 2.2, cw, 0.12, fill=col)
    box(s, x, 2.32, cw, 3.85, fill=GREY)
    text(s, x + 0.22, 2.5, 1.0, 0.6, f"0{i+1}", size=32, color=col if col != CYAN else BLUE, font=LIGHT)
    text(s, x + 0.22, 3.2, cw - 0.4, 0.75, head, size=17, color=NAVY, font=SEMI, line=1.05)
    text(s, x + 0.22, 4.0, cw - 0.4, 2.05, body, size=14, color=TEXT, line=1.15)

# 10 — Forecasting --------------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="WTA IN ACTION", title="Top-down sets ambition. Bottoms-up delivers it.",
              notes=cue(4) + """
Introduce the more sophisticated way WTA will report growth and forecast targets.

KEY LINE: "Top-down sets our annual ambition. Bottoms-up is how you drive it month by month. Eric's team needs your ground-level visibility into pipeline likelihoods, renewal dates and scope expansions to deploy agency resources accurately."

Why it matters: accurate forecasts drive hiring and freelance approvals, resource allocation and whether the agency hits its profit target. A forecast that only changes at quarter-end is a surprise, not a forecast.

The hybrid: leadership sets FY2027 targets by portfolio; account leads own the client-level forecast inside it and manage it dynamically as accounts move.
(Confirm with Eric: WTA's forecast cadence and the tools account leads will update.)
""")
cols = [
    ("TOP-DOWN STRATEGY", "Leadership & Finance", ["Annual growth ambition", "Agency profit targets",
                                                 "FY2027 targets set by portfolio"], NAVY, WHITE),
    ("BOTTOMS-UP REALITY", "Account leads", ["Pipeline likelihood by client", "Renewal dates and risks",
                                            "Scope expansions and reductions", "Monthly forecast updates"], GREY, NAVY),
]
for i, (head, who, items, fill, fg) in enumerate(cols):
    x = 0.52 + i * 6.29
    box(s, x, 2.2, 6.0, 2.95, fill=fill)
    text(s, x + 0.35, 2.4, 5.3, 0.35, head, size=14, color=TEAL if fill == NAVY else RED, font=SEMI)
    text(s, x + 0.35, 2.8, 5.3, 0.45, who, size=22, color=fg, font=BODY)
    text(s, x + 0.35, 3.45, 5.3, 1.6, [{"t": t, "bullet": True, "after": 4,
                                         "bucolor": TEAL if fill == NAVY else RED} for t in items],
         size=15, color=fg)
box(s, 0.52, 5.4, 12.29, 0.85, fill=None, line=NAVY, line_w=1.25)
text(s, 0.85, 5.4, 11.7, 0.85, [
    {"t": [{"t": "The hybrid  ", "font": SEMI, "color": RED},
           {"t": "FY2027 targets are assigned by portfolio and managed dynamically, month by month, by the accounts within it."}]},
], size=16, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)

# 11 — Portfolio buckets --------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="WTA IN ACTION", title="Segment your FY2027 portfolio", notes=cue(4) + """
Every account in your portfolio gets a designation — and each designation has a different commercial posture and margin expectation.

• Growth accounts: proactive upsells; introduce new solution sets.
• Maintenance accounts: standardize scope, drive operational efficiency, defend margin.
• Attrition / declining accounts: manage resource reallocation deliberately to protect margin — move people before the revenue disappears, not after.
• New business: Standardized Solution Sets from proposal day one.

Exercise to suggest for after the session: each account lead tags every client in their portfolio with one of these four before FY2027 planning.
""")
quads = [
    ("Growth", "Proactive upsells; introduce new solution sets", RED),
    ("Maintenance", "Scope standardization, operational efficiency, margin defense", NAVY),
    ("Attrition / declining", "Managed resource reallocation to protect margin", BLUE),
    ("New business", "Standardized Solution Sets from proposal day one", ORANGE),
]
qw, qh = 6.0, 1.9
for i, (head, body, col) in enumerate(quads):
    x = 0.52 + (i % 2) * (qw + 0.29)
    y = 2.2 + (i // 2) * (qh + 0.2)
    box(s, x, y, qw, qh, fill=GREY)
    num_circle(s, x + 0.35, y + 0.35, 0.6, f"0{i+1}", fill=col, size=14)
    text(s, x + 1.2, y + 0.33, qw - 1.5, 0.6, head, size=24, color=NAVY, font=BODY, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + 1.2, y + 1.0, qw - 1.5, 0.75, body, size=15, color=TEXT, line=1.15)


# =====================================================================================
# PART 3 — THE COMMERCIAL OPERATING SYSTEM (38:00 – 55:00)
# =====================================================================================
divider("PART 3  ·  38:00 – 55:00", "The Commercial Operating System", "Compete on value, not hours", """
"Everything so far has been about managing margin inside today's model. Now let's change the model itself."
""", 0.5)

# 12a — Five pillars ------------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="THE COMMERCIAL OPERATING SYSTEM", title="Five pillars, one system",
              notes=cue(1.5) + """
Define the term before we go deeper. A Commercial Operating System is the framework through which an agency consistently creates, prices, sells, contracts, delivers and measures value.

Most agencies have invested heavily in how they deliver client work. Far fewer have invested in how they sell, package, price, contract and govern that work. The result is a commercial model that evolves organically: services get more customized, pricing varies client to client, proposals take longer, and profitability depends on individual judgment.

KEY LINE: "The shift is from individual pricing judgment to a repeatable, governed commercial system — one that aligns strategy, pricing, finance, growth, account leadership and operations."

Keep this brisk; the next slides show what it looks like in practice.
""")
text(s, 0.52, 1.95, 12.29, 0.4,
     "How WTA will consistently create, price, sell, contract, deliver and measure value",
     size=16, color=MUTED)
pillars = [
    ("Commercial strategy", "Where WTA creates differentiated value, the clients it serves best, and how to position it"),
    ("Product architecture", "Expertise turned into defined products with standardized scope, pricing logic and contract-ready language"),
    ("Pricing & revenue model", "A modern pricing methodology that supports profitable growth and simplifies decisions"),
    ("Organizational readiness", "Growth, Strategy, Account Leadership, Operations and Finance able to sell and deliver it"),
    ("Commercial governance", "Reporting, decision frameworks and dashboards that keep the system evolving"),
]
for i, (head, body) in enumerate(pillars):
    y = 2.55 + i * 0.75
    text(s, 0.52, y, 0.9, 0.62, f"0{i+1}", size=30, color=RED, font=LIGHT, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 1.45, y, 3.6, 0.62, head, size=20, color=NAVY, font=SEMI, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 5.15, y, 7.66, 0.62, body, size=14, color=TEXT, anchor=MSO_ANCHOR.MIDDLE, line=1.1)
    rule(s, 0.52, y + 0.69, 12.81, y + 0.69, color=RULE)

# 12 — Products not headcount ---------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="THE COMMERCIAL OPERATING SYSTEM", title="Sell products, not headcount",
              notes=cue(2.5) + """
Today WTA largely sells time and materials and retainers. The evolved model sells defined products.

Old model: "Here is a grid of 17 FTEs across 5 departments at a $190 blended rate."
New model: "Here are three standardized, modular Solution Sets with defined outcomes and fixed pricing."

CRITICAL CLARIFICATION — say all three explicitly; this is where people get confused:
• It is NOT pay-for-performance or contingency risk.
• It is NOT an unbundled, discounted à la carte menu.
• It IS pre-packaged agency expertise with explicit boundaries.

We can still customize to each client's needs — we just don't reinvent the wheel every time.
""")
box(s, 0.52, 2.2, 6.0, 2.2, fill=GREY)
text(s, 0.85, 2.4, 5.4, 0.35, "OLD MODEL  ·  SELLING INPUTS", size=14, color=MUTED, font=SEMI)
text(s, 0.85, 2.85, 5.4, 1.4, "“Here is a grid of 17 FTEs across 5 departments at a $190 blended rate.”",
     size=20, color=TEXT, font=LIGHT, line=1.1)
box(s, 6.81, 2.2, 6.0, 2.2, fill=NAVY)
text(s, 7.14, 2.4, 5.4, 0.35, "NEW MODEL  ·  PRICING IMPACT", size=14, color=TEAL, font=SEMI)
text(s, 7.14, 2.85, 5.4, 1.4,
     "“Here are three standardized, modular Solution Sets with defined outcomes and fixed pricing.”",
     size=20, color=WHITE, font=BODY, line=1.1)
chips = [("✕", "It is NOT", "pay-for-performance or contingency risk", MUTED),
         ("✕", "It is NOT", "an unbundled, discounted à la carte menu", MUTED),
         ("✓", "It IS", "pre-packaged agency expertise with explicit boundaries", RED)]
cw, gap = 3.93, 0.25
for i, (mark, lead, body, col) in enumerate(chips):
    x = 0.52 + i * (cw + gap)
    box(s, x, 4.7, cw, 1.5, fill=None, line=col, line_w=1.25)
    num_circle(s, x + 0.3, 4.98, 0.5, mark, fill=col, size=16)
    text(s, x + 1.0, 4.85, cw - 1.2, 1.25, [
        {"t": lead, "font": SEMI, "color": col if col == RED else NAVY, "size": 16, "after": 2},
        {"t": body}], size=14, color=TEXT, line=1.1)

# 13 — Four-step framework ------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="THE COMMERCIAL OPERATING SYSTEM", title="The four-step productization framework",
              notes=cue(3.5) + """
1. Define the Unit — what is the client actually buying? A campaign, a platform, a monthly stewardship. Clients don't care about 25% of an ACD, FTE counts, utilization rates or our staffing plan. They care about what their investment delivers.
2. Define the Product — document exact inclusions and explicit exclusions: volume, platforms, asset types, cadence, rounds of revision, client responsibilities. "Ambiguity is where margin bleeds." Every undefined element is a future scope dispute.
3. Price the Product — use historical effort to inform cost, but set the price on market value and client impact. Time informs pricing; it does not define it.
4. Build the Offering — clear, decision-ready, menu-style proposals: product, scope definition, price per unit, quantity. Everything the client needs to decide in one view.

Who participates: account leadership shapes the product from the client's perspective; commercial finance ensures it's clear, complete, priced consistently and protects margin.
""")
steps = [
    ("Define the unit", "Identify what the client is actually buying: a campaign, a platform, a monthly stewardship."),
    ("Define the product", "Document exact inclusions and explicit exclusions."),
    ("Price the product", "Historical effort informs cost; market value and client impact set price."),
    ("Build the offering", "Present clear, decision-ready, menu-style proposals."),
]
sw = 2.9
xs = [0.52 + i * (sw + 0.23) for i in range(4)]
rule(s, xs[0] + 0.4, 2.65, xs[-1] + 0.4, 2.65, color=RULE, width=1.5)
for i, (head, body) in enumerate(steps):
    x = xs[i]
    num_circle(s, x, 2.25, 0.8, f"0{i+1}", fill=NAVY if i < 3 else RED, size=18)
    text(s, x, 3.3, sw, 0.45, head, size=20, color=NAVY, font=SEMI)
    text(s, x, 3.8, sw - 0.15, 1.4, body, size=15, color=TEXT, line=1.15)
q = box(s, 0.52, 5.35, 12.29, 0.9, fill=GREY)
text(s, 0.85, 5.35, 11.7, 0.9, [
    {"t": [{"t": "Ambiguity is where margin bleeds. ", "font": SEMI, "color": RED},
           {"t": "Every undefined element is a future scope dispute."}]},
], size=18, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)

# 14 — Case study: before -------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="CASE STUDY  ·  THE RETAIL PITCH WIN", title="How it started: three staffing grids",
              notes=cue(2.5) + """
Illustrative retail parallel (coordinate the WTA framing with Abby before the session). Keep the prospect and agency unnamed.

The setup: a large retail prospect issued a complex, multi-year RFP — no stated budget, shifting priorities, multiple initiatives, and a strong incumbent with deep relationships. On paper, unwinnable.

The first response was the classic one: three FTE-based staffing options at a $190 blended hourly rate, over 19 months — $7.4M, $8.4M and $10.1M; 17 to 26 FTEs.

The trap: no defined outcomes, no scope guardrails, no way for the client to know what they were actually buying. Procurement got bogged down auditing hours. Pricing drove to hours, not value.
""")
opts = [("ESSENTIAL", "17", "48,450", "$7.4M"),
        ("SIGNATURE", "21", "59,850", "$8.4M"),
        ("FULL-SCALE", "26", "74,100", "$10.1M")]
cw, gap = 3.93, 0.25
for i, (name, fte, hrs, fee) in enumerate(opts):
    x = 0.52 + i * (cw + gap)
    box(s, x, 2.2, cw, 2.75, fill=GREY)
    text(s, x + 0.3, 2.38, cw - 0.6, 0.35, f"{name} OPTION", size=14, color=MUTED, font=SEMI)
    text(s, x + 0.3, 2.8, cw - 0.6, 0.9, fee, size=48, color=NAVY, font=LIGHT)
    rule(s, x + 0.3, 3.85, x + cw - 0.3, 3.85, color=RULE)
    text(s, x + 0.3, 4.0, cw - 0.6, 0.8, [
        {"t": f"{fte} FTEs  ·  {hrs} hours", "font": SEMI, "color": NAVY},
        {"t": "$190 blended rate  ·  19 months", "color": MUTED}], size=14, line=1.15)
box(s, 0.52, 5.2, 12.29, 1.05, fill=None, line=RED, line_w=1.25)
text(s, 0.85, 5.2, 11.7, 1.05, [
    {"t": [{"t": "The trap  ", "font": SEMI, "color": RED},
           {"t": "No defined outcomes, no scope guardrails, no way for the client to know what they were buying. Procurement bogged down in hourly audits."}]},
], size=16, color=NAVY, anchor=MSO_ANCHOR.MIDDLE, line=1.1)

# 15 — Case study: after --------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="CASE STUDY  ·  THE RETAIL PITCH WIN", title="The productized breakthrough",
              notes=cue(3) + """
The rework: three products, each independently purchasable. The client builds its own portfolio.

• Pulse Campaigns — targeted digital activation across 2 platforms, $75,000 per campaign. The client dials volume up or down; zero hourly renegotiation.
• Milestone Activations — full multi-channel brand moments with full creative production, $500,000 per activation.
• Monthly Strategic Stewardship — ongoing continuity, reporting and strategic advisory, $25,000 per month. No separate SOW required.

The outcome: the underdog won — displacing a long-standing incumbent.

KEY LINE: "The creative was great, but the commercial model was the deciding factor. The client had total budget predictability, and the agency had zero staffing debates."
""")
prods = [("Pulse Campaigns", "$75K", "per campaign",
          "Targeted 2-platform digital activations. Client dials volume up or down with zero hourly renegotiation."),
         ("Milestone Activations", "$500K", "per activation",
          "Full multi-channel brand moments, full creative production included."),
         ("Strategic Stewardship", "$25K", "per month",
          "Ongoing continuity, reporting and strategic advisory.")]
cw, gap = 3.93, 0.25
for i, (name, price, unit, body) in enumerate(prods):
    x = 0.52 + i * (cw + gap)
    box(s, x, 2.2, cw, 0.5, fill=[NAVY, BLUE, RED][i])
    text(s, x + 0.3, 2.2, cw - 0.6, 0.5, name.upper(), size=14, color=WHITE, font=SEMI, anchor=MSO_ANCHOR.MIDDLE)
    box(s, x, 2.7, cw, 2.35, fill=GREY)
    text(s, x + 0.3, 2.85, cw - 0.6, 0.9, [
        {"t": [{"t": price, "size": 44, "font": LIGHT, "color": NAVY},
               {"t": "  " + unit, "size": 14, "color": MUTED}]}])
    text(s, x + 0.3, 3.85, cw - 0.6, 1.1, body, size=14, color=TEXT, line=1.15)
box(s, 0.52, 5.3, 12.29, 0.95, fill=NAVY)
text(s, 0.85, 5.3, 11.7, 0.95, [
    {"t": [{"t": "The underdog won.  ", "font": SEMI, "color": WHITE},
           {"t": "Creative mattered, but the commercial model was the deciding factor.", "font": LIGHT}]},
], size=20, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

# 16 — What we gain -------------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="THE COMMERCIAL OPERATING SYSTEM", title="What we gain: speed, margin and sanity",
              notes=cue(1.5) + """
For the client: total clarity on what they're buying, flexibility to adjust volume without re-contracting, and budget certainty.

For WTA account leads:
• Faster speed to sale — no custom staffing plan built from scratch for every proposal.
• No timesheet arguments with clients — we talk about outcomes, not hours.
• Clear change orders — inclusions and exclusions are already written down, so out-of-scope is obvious to everyone.
• Delivery teams execute proven solution sets instead of reinventing the wheel; repeating the same products builds deeper expertise, and there's no translation gap between what's in the contract and what gets executed.

Bottom row is a PLACEHOLDER for a supporting case study and its results. Add or remove before the session.
""")
cols = [
    ("FOR THE CLIENT", ["Total clarity on deliverables", "Flexibility without re-contracting", "Budget certainty"], GREY, NAVY, RED),
    ("FOR WTA ACCOUNT LEADS", ["Faster speed to sale", "No timesheet arguments with clients", "Clear, simple change orders",
                               "Delivery runs proven solution sets"], NAVY, WHITE, TEAL),
]
for i, (head, items, fill, fg, acc) in enumerate(cols):
    x = 0.52 + i * 6.29
    box(s, x, 2.2, 6.0, 2.45, fill=fill)
    text(s, x + 0.35, 2.38, 5.3, 0.35, head, size=14, color=acc, font=SEMI)
    text(s, x + 0.35, 2.85, 5.3, 1.75, [{"t": t, "bullet": True, "after": 3, "bucolor": acc} for t in items],
         size=16, color=fg)
box(s, 0.52, 4.9, 12.29, 1.35, fill=None, line=MUTED, line_w=1)
text(s, 0.85, 4.9, 11.7, 1.35, [
    {"t": "CASE STUDY PLACEHOLDER", "font": SEMI, "color": MUTED, "size": 12, "after": 4},
    {"t": "Supporting case study and results to be added", "color": MUTED, "size": 18, "font": LIGHT}],
    anchor=MSO_ANCHOR.MIDDLE)

# 17 — Rollout ------------------------------------------------------------------------
s = new_slide(L_CONTENT, eyebrow="THE COMMERCIAL OPERATING SYSTEM", title="The WTA phased rollout", notes=cue(2) + """
We are starting with new business, and phasing the model into existing clients deliberately — client by client, with a strategic plan for when and how to introduce it.

• Phase 1 (immediate / Q4): inbound RFPs and new business lead with Solution Sets.
• Phase 2 (FY2027 planning): identify which existing client contracts transition during their annual renewal cycles, with a plan per client.
• Phase 3 (January 1, 2027): full adoption across budgeting, forecasting and portfolio management.

ENABLEMENT: reassure the room that account leads will be supported through implementation, so no one has to invent pricing from scratch.

FINANCIAL GUARDRAIL (Eric's domain): timesheets are still required internally as the legal and audit record that measures direct labor cost and margin. We simply stop using them as the public pricing mechanism for clients.
""")
phases = [("PHASE 1", "Immediate / Q4", "Inbound RFPs and new business lead with Solution Sets", NAVY),
          ("PHASE 2", "FY2027 planning", "Identify existing client contracts to transition at annual renewal, with a plan per client", BLUE),
          ("PHASE 3", "January 1, 2027", "Full adoption across budgeting, forecasting and portfolio management", RED)]
pw, gap = 3.93, 0.25
for i, (ph_, when, body, col) in enumerate(phases):
    x = 0.52 + i * (pw + gap)
    box(s, x, 2.2, pw, 0.12, fill=col)
    box(s, x, 2.32, pw, 2.6, fill=GREY)
    text(s, x + 0.3, 2.5, pw - 0.6, 0.35, ph_, size=14, color=col, font=SEMI)
    text(s, x + 0.3, 2.85, pw - 0.6, 0.45, when, size=20, color=NAVY, font=SEMI)
    text(s, x + 0.3, 3.4, pw - 0.6, 1.3, body, size=14, color=TEXT, line=1.15)
box(s, 0.52, 5.2, 12.29, 1.05, fill=NAVY)
text(s, 0.85, 5.2, 11.7, 1.05, [
    {"t": [{"t": "Where we start  ", "font": SEMI, "color": TEAL},
           {"t": "New business first. Existing clients follow, each with a plan for when and how the model is introduced."}]},
], size=16, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, line=1.1)


# =====================================================================================
# THE WRAP (55:00 – 60:00)
# =====================================================================================

# 18 — Golden rules -------------------------------------------------------------------
s = new_slide(L_DARK, eyebrow="THE COMMERCIAL GOLDEN RULES", notes=cue(2) + """
If you remember nothing else from today, remember these three.

1. Direct Margin (55%–62%) is your primary health metric.
2. Time informs our cost — it does not define our price. (Timesheets stay: they're how we measure cost and margin internally.)
3. Just because our team can do something doesn't mean we should sell it without a defined scope. Productization creates focus; focus creates margin.
""")
rules_ = [("Direct Margin of 55–62%", "is your primary health metric."),
          ("Time informs our cost.", "It does not define our price."),
          ("Just because we can doesn’t mean we should", "sell it without a defined scope.")]
for i, (a_, b_) in enumerate(rules_):
    y = 1.35 + i * 1.65
    text(s, 0.52, y, 1.4, 1.2, f"0{i+1}", size=60, color=WHITE, font=LIGHT, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 2.1, y, 10.5, 1.2, [
        {"t": a_, "font": SEMI, "size": 30, "color": WHITE, "after": 2},
        {"t": b_, "font": LIGHT, "size": 26, "color": WHITE}], anchor=MSO_ANCHOR.MIDDLE, line=1.0)
    if i < 2:
        rule(s, 2.1, y + 1.42, 12.81, y + 1.42, color=WHITE, width=0.5)

# 19 — Q&A ----------------------------------------------------------------------------
s = new_slide(L_CLOSING, title="Let’s build FY2027 together", keep=(16,), notes=cue(3) + """
Closing line before handing over: "Agencies shouldn't compete on how many hours they sell. They should compete on how effectively they create value — and how intentionally they capture it."

Invite Eric (CFO) and commercial leadership to share brief closing remarks.
Then open the floor to account leads for tactical questions.

Likely questions to be ready for:
• "Do we still fill out timesheets?" — Yes. Internally they measure cost and margin and are our legal/audit record; they're just no longer the client's price tag.
• "What happens to my existing retainer clients?" — Nothing changes overnight. Phase 2 builds a client-by-client transition plan around renewal dates.
• "What if a client insists on hours?" — We can still show effort internally to inform cost; the conversation we lead with is products and outcomes.
• "Who sets the prices?" — Account leadership and commercial finance together, using shared pricing tools and a common product catalog.
""")
t = ph(s, 0)
t.left, t.top, t.width, t.height = Inches(0.52), Inches(2.2), Inches(12.29), Inches(1.2)
for r in t.text_frame.paragraphs[0].runs:
    r.font.size = Pt(54)
b = ph(s, 16)
b.text_frame.text = ""
tf = b.text_frame
for i, (lead, rest) in enumerate([("Closing remarks  ", "Eric, CFO, and commercial leadership"),
                                  ("Open floor  ", "Your tactical questions")]):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = lead; _style_run(r, 18, WHITE, SEMI)
    r = p.add_run(); r.text = rest; _style_run(r, 18, WHITE, BODY)

# Sections for easy navigation in PowerPoint ---------------------------------------
SECTIONS = [("The Provocation", 0), ("Part 1 – Agency Economics 101", 3), ("Part 2 – WTA in Action", 8),
            ("Part 3 – The Commercial Operating System", 13), ("The Wrap & Q&A", 21)]
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

prs.core_properties.title = "Stop Selling Hours: The WTA Commercial Operating System"
prs.core_properties.author = "Erin Jemas, Price It Right"
prs.core_properties.subject = "Maple Ranch account leadership session"
prs.save(str(OUT))
print(f"Saved {OUT}  ({len(ids)} slides, timed to {_clock[0]:g} min)")
