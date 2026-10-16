"""Build the Confidence 2026 Year-End All Hands planning proposal on the VML 2026 template.

Usage: python3 build_deck.py [output.pptx]

Every slide sits on a native VML layout (light and dark masters, header rules, spark mark,
slide number), and all type uses the template's Snowflake Sans theme fonts. Color follows the
brand rule: black and white first, with one supporting hue (Gold) used sparingly.
"""
import copy
import sys
import uuid
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.shapes import PP_PLACEHOLDER as PH
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE.parent / "template" / "VML-Template-Feb-2026.pptx"
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE.parent / "Confidence-2026-All-Hands-Planning-Proposal.pptx"

# --- VML palette (template "VML" scheme + brand grayscale + one secondary hue: Gold) ----
BLACK = RGBColor(0x19, 0x19, 0x19)
WHITE = RGBColor(0xF5, 0xF5, 0xF5)
DGRAY = RGBColor(0x36, 0x37, 0x32)
MGRAY = RGBColor(0x53, 0x54, 0x4A)
GRAY = RGBColor(0x8B, 0x8C, 0x81)
LGRAY = RGBColor(0xC2, 0xC4, 0xB8)
GOLD = RGBColor(0xD3, 0xA3, 0x4A)
LGOLD = RGBColor(0xF9, 0xD2, 0x8C)
DGOLD = RGBColor(0xB3, 0x7A, 0x00)

BODY = "+mn-lt"   # Snowflake Sans Book (theme minor font)
HEAD = "+mj-lt"   # Snowflake Sans Medium (theme major font)

# Content grid (inches), matching the template's header rules at 0.55" and 1.61"
L, R = 0.29, 13.05
W = R - L
TOP = 1.95
BOTTOM = 7.05

prs = Presentation(str(TEMPLATE))
LIGHT, DARK = prs.slide_masters[0], prs.slide_masters[1]


def layout(master, name):
    for lay in master.slide_layouts:
        if lay.name == name:
            return lay
    raise KeyError(name)


# --- remove the template's guidance and sample slides ----------------------------------
sldIdLst = prs.slides._sldIdLst
for sldId in list(sldIdLst):
    prs.part.drop_rel(sldId.rId)
    sldIdLst.remove(sldId)


# --- helpers ---------------------------------------------------------------------------
def _inherit_slide_number(slide, lay):
    """python-pptx skips slide-number placeholders; clone it from the layout."""
    for ph in lay.placeholders:
        if ph.placeholder_format.type == PH.SLIDE_NUMBER:
            slide.shapes._spTree.append(copy.deepcopy(ph._element))


def new_slide(master, name, title=None, notes=None, keep=()):
    lay = layout(master, name)
    s = prs.slides.add_slide(lay)
    for ph in list(s.placeholders):
        idx = ph.placeholder_format.idx
        if ph.placeholder_format.type in (PH.TITLE, PH.CENTER_TITLE) and title is not None:
            ph.text_frame.text = title
        elif idx in keep:
            continue
        else:
            ph._element.getparent().remove(ph._element)
    _inherit_slide_number(s, lay)
    if notes:
        s.notes_slide.notes_text_frame.text = notes.strip()
    return s


def content(title, notes, dark=False):
    """Interior 1.0: small header between the two template rules, open content area below."""
    return new_slide(DARK if dark else LIGHT, f"Interior 1.0 ({'Dark' if dark else 'Light'})",
                     title=title, notes=notes)


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


def text(slide, x, y, w, h, paras, size=14, color=BLACK, font=BODY, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, after=0, line=None, name=None):
    """paras: str | list of str | list of dict(t=..., size, color, font, bold, italic, align, after, check)
    A dict's t may itself be a list of run dicts for mixed styling. check=True adds the VML □ bullet."""
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
        if p.get("check"):
            pPr = para._p.get_or_add_pPr()
            pPr.set("marL", str(int(Inches(0.3))))
            pPr.set("indent", str(-int(Inches(0.3))))
            etree.SubElement(pPr, qn("a:buFont"), typeface=BODY)
            etree.SubElement(pPr, qn("a:buChar"), char="□")
    return tb


def box(slide, x, y, w, h, fill=None, line=None, line_w=0.75, name=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
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
    s.text_frame.text = ""
    return s


def rule(slide, x1, y1, x2, y2, color=LGRAY, width=0.75):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = color
    c.line.width = Pt(width)
    return c


def label_in(shape, t, size=12, color=WHITE, font=HEAD, align=PP_ALIGN.CENTER):
    tf = shape.text_frame
    tf.margin_left = tf.margin_right = Inches(0.06)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = t
    _style_run(r, size, color, font)


def table(slide, x, y, w, col_w, rows, row_h=0.5, size=12, head_color=BLACK, body_color=BLACK,
          rule_color=LGRAY, highlight=None):
    """Editable native table in the VML style: no fills, hairline row rules, bold header row."""
    gf = slide.shapes.add_table(len(rows), len(col_w), Inches(x), Inches(y), Inches(w), Inches(row_h * len(rows)))
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    for attr in ("firstRow", "bandRow"):
        tblPr.set(attr, "0")
    style = tblPr.find(qn("a:tableStyleId"))
    if style is None:
        style = etree.SubElement(tblPr, qn("a:tableStyleId"))
    style.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # No Style, No Grid
    for ci, cw in enumerate(col_w):
        tbl.columns[ci].width = Inches(cw)
    for ri, row in enumerate(rows):
        tbl.rows[ri].height = Inches(row_h)
        for ci, val in enumerate(row):
            c = tbl.cell(ri, ci)
            if highlight is not None and ri == highlight:
                c.fill.solid()
                c.fill.fore_color.rgb = LGOLD
            else:
                c.fill.background()
            c.margin_left = Inches(0.08 if ci else 0)
            c.margin_right = Inches(0.08)
            c.margin_top = c.margin_bottom = Inches(0.04)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = c.text_frame
            tf.word_wrap = True
            tf.text = ""
            r = tf.paragraphs[0].add_run()
            r.text = val
            head = ri == 0
            _style_run(r, size, head_color if head else body_color, HEAD if (head or ci == 0) else BODY)
            tcPr = c._tc.get_or_add_tcPr()
            for side in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                ln = etree.SubElement(tcPr, qn(side), w=str(int(Pt(1 if (head and side == "a:lnB") else 0.75))))
                if side == "a:lnB":
                    sf = etree.SubElement(ln, qn("a:solidFill"))
                    etree.SubElement(sf, qn("a:srgbClr"), val=str(head_color if head else rule_color))
                else:
                    etree.SubElement(ln, qn("a:noFill"))
            # schema order: borders must precede the cell fill
            for fill_tag in ("a:solidFill", "a:noFill"):
                for el in tcPr.findall(qn(fill_tag)):
                    tcPr.remove(el)
                    tcPr.append(el)
    return gf


def numeral(slide, x, y, n, color=BLACK, size=40):
    return text(slide, x, y, 2.2, 0.7 if size < 60 else 1.2, f"{n:02d}", size=size, color=color, font=HEAD)


# =======================================================================================
# 1 — COVER
# =======================================================================================
s = new_slide(DARK, "Cover 1.0 (Dark)", keep=(0, 13, 14, 15), notes="""
Purpose of this session: align Confidence leadership, group leads and pod leads on the concept for our final All Hands of 2026, and make the decisions that let planning start next week.

Travis Kubicek is proposed as planning lead (to confirm), with Zack Jemas and the group and pod leads supporting.

This is a working draft. The goal today is alignment on the concept, not a detailed content review.
""")
t = ph(s, 0)
t.text_frame.text = "Confidence 2026:\nOur Impact,\nOur Future"
for p in t.text_frame.paragraphs:
    for r in p.runs:
        r.font.size = Pt(80)
text(s, L, 5.35, 9.0, 0.5, "Year-End All Hands  |  Planning proposal for Confidence Leads", size=20, color=GOLD)
for idx, val in ((13, "Oct 2026"), (14, "Working draft"), (15, "Target: Nov 23 or 24, 2026")):
    ph(s, idx).text_frame.text = val

# =======================================================================================
# 2 — THE IDEA
# =======================================================================================
s = content("The proposal: a 2026 Impact Report and 2027 Future Outlook, not a traditional celebration", """
The one-line version: rather than a traditional year-end celebration, our last All Hands of the year becomes a 2026 Impact Report and a 2027 Future Outlook.

It is organized around three themes:
• Our Impact: the breadth, scale and value of what Confidence contributed across clients, disciplines and capabilities.
• Our People: the individuals and teams behind that work, including how we grew, collaborated and supported each other through a challenging year.
• Our Future: emerging trends inspired by VML Intelligence's Future 100: 2026, plus an imaginative, cross-practice conversation.

About 10 minutes stays reserved for Confidence leadership updates.

The intent is an experience that feels true to the year we've had, reinforces the value of our collective work, and gives people room to think beyond their immediate projects.
""")
text(s, L, TOP, W, 0.6, [{"t": [
    {"t": "Recognize our impact.  ", "color": BLACK},
    {"t": "Appreciate our people.  ", "color": BLACK},
    {"t": "Imagine our future.", "color": DGOLD},
]}], size=26, font=HEAD)
cols = [
    ("Our Impact", "15 min",
     "Show the breadth, scale and value of what Confidence delivered across clients, disciplines and capabilities."),
    ("Our People", "8 min",
     "Recognize the people behind the work: how we grew, collaborated and supported one another through a hard year."),
    ("Our Future", "17 min",
     "Three Future 100-inspired predictions, then a live, cross-practice “yes, and” conversation about 2027."),
]
cw, gap, y0, ch = (W - 2 * 0.3) / 3, 0.3, 3.0, 2.85
for i, (head, mins, body) in enumerate(cols):
    x = L + i * (cw + gap)
    box(s, x, y0, cw, ch, line=BLACK)
    numeral(s, x + 0.25, y0 + 0.2, i + 1, color=DGOLD if i == 2 else BLACK)
    text(s, x + cw - 1.45, y0 + 0.32, 1.2, 0.4, mins, size=14, color=MGRAY, align=PP_ALIGN.RIGHT)
    text(s, x + 0.25, y0 + 1.0, cw - 0.5, 0.45, head, size=22, font=HEAD)
    text(s, x + 0.25, y0 + 1.55, cw - 0.5, 1.2, body, size=14, line=1.1)
text(s, L, 6.2, W, 0.8, [
    {"t": [{"t": "Plus 10 min of Confidence leadership updates and a 5-min open and close.  ", "font": HEAD},
           {"t": "60 minutes total. Target Nov 23 or 24, 2026; fallback first week of December."}]},
], size=14, color=MGRAY)

# =======================================================================================
# 3 — TODAY'S ASK
# =======================================================================================
s = content("What we need from this group today", """
We don't need to walk every section of the plan today. The full detail is in the deck and appendix for reference.

Focus the conversation on three decisions:
1. Does the overall concept feel right? Impact, People, Future, in a 60-minute format, before Thanksgiving.
2. Which stories best represent our practice? We're proposing United Rentals, Sherwin-Williams and Mazda as featured stories, with a broader set of highlights.
3. How ambitious should we be with the collaborative future discussion? A live popcorn conversation fits 60 minutes; breakouts would need about 70.

If we get through those, confirm the practical pieces: Travis as planning lead, a small planning group, and an AQ host.
""", dark=True)
asks = [
    ("Does the concept feel right?",
     "Impact, People and Future, in 60 minutes, shortly after November reporting and before Thanksgiving."),
    ("Which stories best represent our practice?",
     "United Rentals, Sherwin-Williams and Mazda as featured stories, plus six to eight highlights from across the practice."),
    ("How ambitious should the future discussion be?",
     "A live popcorn conversation fits 60 minutes. Small-group breakouts would need about 70."),
]
cw = (W - 2 * 0.4) / 3
for i, (q, d) in enumerate(asks):
    x = L + i * (cw + 0.4)
    numeral(s, x, TOP + 0.1, i + 1, color=GOLD, size=72)
    rule(s, x, 3.35, x + cw, 3.35, color=MGRAY)
    text(s, x, 3.55, cw - 0.2, 1.0, q, size=22, color=WHITE, font=HEAD, line=1.05)
    text(s, x, 4.85, cw - 0.2, 1.3, d, size=14, color=LGRAY, line=1.15)
text(s, L, 6.45, W, 0.5, [{"t": [
    {"t": "Then confirm the practical pieces:  ", "font": HEAD, "color": WHITE},
    {"t": "planning lead (Travis), a small planning group, an AQ host, and how we collect impact data and nominations.",
     "color": LGRAY},
]}], size=14)

# =======================================================================================
# 4 — WHY THIS FORMAT
# =======================================================================================
s = content("Why this format: the year we've had calls for impact and perspective, not just celebration", """
Leadership direction: in planning conversations, Kylie asked us to adjust the traditional celebration format to reflect the current environment. There's plenty to recognize, but an overly celebratory tone may feel disconnected from the challenges, uncertainty and demands people have experienced this year. She also asked us to make room for the reality of people's experiences while still creating something productive and forward-looking.

The evidence: a review of the February, April, June, August and October Technology Reports shows five recurring themes (on the right).

Together they suggest one story: 2026 was a year of evolving from individual project successes toward reusable capabilities that strengthen delivery across Confidence.

That sets up the bridge to the future discussion: if our work, capabilities and roles have already evolved this much, what could Confidence look like in 2027? This framing connects the retrospective and forward-looking halves so they don't feel like unrelated presentations.
""")
text(s, L, TOP, 4.3, 0.35, "What Kylie asked for", size=16, font=HEAD)
text(s, L, TOP + 0.5, 4.3, 3.3, [
    {"t": "An impact-focused retrospective, not an overly celebratory one", "check": True},
    {"t": "Recognition tied to actual contributions", "check": True},
    {"t": "Inspiration from VML's Future 100", "check": True},
    {"t": "A popcorn-style conversation that builds on ideas", "check": True},
    {"t": "Room to imagine reinventing the practice, not just improving it", "check": True},
], size=14, after=8, line=1.05)
x2 = 5.0
text(s, x2, TOP, R - x2, 0.35, "What the 2026 Technology Reports tell us", size=16, font=HEAD)
themes = [
    ("AI moved from experimentation into delivery", "Analysis, development, QA, agent governance and reporting"),
    ("Quality engineering became continuous", "More automation, earlier validation, better quality metrics"),
    ("Shared delivery practices became a differentiator", "Common Agile frameworks, templates, onboarding, governance"),
    ("We grew into a strategic partner", "Shaping solutions, managing risk, influencing technology decisions"),
    ("Our people delivered through change", "Delays, transitions and shifting priorities, with commitments kept"),
]
for i, (h, d) in enumerate(themes):
    y = TOP + 0.5 + i * 0.6
    text(s, x2, y, 0.5, 0.5, f"{i + 1:02d}", size=14, color=DGOLD, font=HEAD)
    text(s, x2 + 0.5, y, 3.9, 0.55, h, size=14, font=HEAD, line=1.0)
    text(s, x2 + 4.5, y, R - x2 - 4.5, 0.55, d, size=13, color=MGRAY, line=1.0)
    rule(s, x2, y + 0.53, R, y + 0.53)
b = box(s, L, 5.6, W, 1.4, fill=BLACK)
text(s, L + 0.35, 5.78, W - 0.7, 1.1, [
    {"t": "The story: 2026 was the year we evolved from individual project successes toward reusable capabilities "
          "that strengthen delivery across Confidence.", "color": WHITE, "font": HEAD, "size": 16, "after": 6},
    {"t": "The bridge: if our work, capabilities and roles have already evolved this much, what could Confidence look like in 2027?",
     "color": LGOLD, "size": 15},
], line=1.05)

# =======================================================================================
# 5 — RUN OF SHOW
# =======================================================================================
s = content("Run of show: 60 minutes, with most of the time on Impact, People and Future", """
Walk the bar left to right. The proportions are the point: leadership updates keep a dedicated 10 minutes, while most of the hour goes to impact, recognition and the future discussion.

Timing can be adjusted once leadership and presenter needs are confirmed. The one segment we recommend protecting is the 12-minute popcorn conversation (more on that shortly).

If the group wants small-group breakouts instead of a live conversation, the meeting would need to extend to about 70 minutes.
""")
segs = [
    (5, "Welcome and opening", "Set the tone; introduce the three themes", "Host / practice leadership", "frame"),
    (10, "Leadership updates", "Organizational, business and practice updates", "Confidence leadership", "frame"),
    (15, "Our Impact: 2026 in review", "Client outcomes, delivery, practice evolution", "Impact report", "core"),
    (8, "Our People: behind the impact", "Recognition tied to real contributions", "Contributor spotlights", "core"),
    (5, "Looking ahead: the Confidence Future 3", "Three Future 100-inspired predictions", "Future Lab", "core"),
    (12, "Cross-practice popcorn conversation", "Collaborative, imaginative “yes, and” discussion", "Moderator + 4–5 voices", "key"),
    (5, "Closing and looking forward", "Appreciation, takeaways, capturing ideas", "Host / leadership", "frame"),
]
fills = {"frame": (LGRAY, BLACK), "core": (BLACK, WHITE), "key": (GOLD, BLACK)}
unit = W / 60
x, t0 = L, 0
bar_y, bar_h = TOP + 0.05, 0.55
for mins, name, focus, who, kind in segs:
    f, c = fills[kind]
    seg = box(s, x, bar_y, mins * unit - 0.04, bar_h, fill=f)
    label_in(seg, f"{mins} min", size=12, color=c)
    x += mins * unit
for m in (0, 15, 30, 45, 60):
    xm = L + m * unit
    text(s, xm - (0.8 if m == 60 else 0.4) if m else xm, bar_y + bar_h + 0.06, 0.8, 0.25, f"{m}′" if m else "0′", size=10, color=MGRAY,
         align=PP_ALIGN.CENTER if 0 < m < 60 else (PP_ALIGN.LEFT if m == 0 else PP_ALIGN.RIGHT))
ty = 3.3
rows_h = 0.52
text(s, L + 0.45, ty - 0.35, 1.3, 0.3, "Time", size=11, color=MGRAY, font=HEAD)
text(s, L + 1.8, ty - 0.35, 4.0, 0.3, "Segment", size=11, color=MGRAY, font=HEAD)
text(s, L + 6.1, ty - 0.35, 4.0, 0.3, "Focus", size=11, color=MGRAY, font=HEAD)
text(s, L + 10.3, ty - 0.35, 2.4, 0.3, "Who", size=11, color=MGRAY, font=HEAD)
t0 = 0
for i, (mins, name, focus, who, kind) in enumerate(segs):
    y = ty + i * rows_h
    f, _ = fills[kind]
    box(s, L, y + 0.15, 0.2, 0.2, fill=f, line=BLACK if kind == "frame" else None)
    text(s, L + 0.45, y + 0.1, 1.3, 0.35, f"{t0:02d}–{t0 + mins:02d}", size=14, color=MGRAY)
    text(s, L + 1.8, y + 0.1, 4.2, 0.35, name, size=14, font=HEAD)
    text(s, L + 6.1, y + 0.1, 4.1, 0.35, focus, size=13)
    text(s, L + 10.3, y + 0.1, 2.46, 0.35, who, size=13, color=MGRAY)
    rule(s, L, y + rows_h, R, y + rows_h)
    t0 += mins

# =======================================================================================
# 6 — BOOKENDS
# =======================================================================================
s = content("Opening, leadership updates and close: short, intentional and coordinated", """
These three segments frame the content.

Welcome and opening (5 min): acknowledge the year honestly and explain the intention. This isn't only a celebration; it's a chance to appreciate what we accomplished and look creatively at the future. Introduce the three themes.

Leadership updates (10 min): potential topics include organizational and business developments, client portfolio outlook, evolving practice priorities and capabilities, initiatives or changes affecting the team, and expectations heading into 2027. Recommend two or three concise updates rather than several long presentations, and have presenters coordinate topics so nothing repeats and time holds.

Closing (5 min): recap the most compelling takeaways, highlight themes from the popcorn conversation and explain how ideas will be captured. Close with appreciation and perspective, not new assignments.

Decision needed: who presents leadership updates, and which topics need to be covered?
""")
cols = [
    ("Welcome and opening", "5 min  ·  Host and/or practice leadership", [
        "Acknowledge the year honestly",
        "Explain why this meeting looks different",
        "Introduce Impact, People and Future",
    ], "An appreciative, inclusive, forward-looking tone"),
    ("Leadership updates", "10 min  ·  Confidence leadership", [
        "Two or three concise updates, not several long ones",
        "Organization, client outlook, practice priorities, 2027 considerations",
        "Presenters coordinate topics to avoid repeats",
    ], "Transparency and alignment, without crowding out the rest"),
    ("Closing", "5 min  ·  Host / leadership", [
        "Recap the most compelling takeaways",
        "Share themes from the popcorn conversation",
        "Explain how ideas will be captured",
    ], "People leave appreciated, connected and curious"),
]
cw = (W - 2 * 0.3) / 3
for i, (head, meta, items, outcome) in enumerate(cols):
    x = L + i * (cw + 0.3)
    box(s, x, TOP, cw, 4.15, line=BLACK)
    text(s, x + 0.25, TOP + 0.25, cw - 0.5, 0.4, head, size=20, font=HEAD)
    text(s, x + 0.25, TOP + 0.75, cw - 0.5, 0.3, meta, size=12, color=MGRAY)
    text(s, x + 0.25, TOP + 1.3, cw - 0.5, 2.2, [{"t": it, "check": True} for it in items], size=14, after=8, line=1.05)
    rule(s, x + 0.25, TOP + 3.25, x + cw - 0.25, TOP + 3.25)
    text(s, x + 0.25, TOP + 3.38, cw - 0.5, 0.7, [{"t": [
        {"t": "Outcome: ", "font": HEAD}, {"t": outcome}]}], size=12, line=1.05)
text(s, L, 6.4, W, 0.5, [{"t": [
    {"t": "Decision needed:  ", "font": HEAD, "color": DGOLD},
    {"t": "who presents the leadership updates, and which topics they need to cover."},
]}], size=14)

# =======================================================================================
# 7 — OUR IMPACT: FEATURED STORIES
# =======================================================================================
s = content("Our Impact (15 min): three featured stories show how the practice evolved", """
Organize the impact report around themes rather than by client or reporting group. Suggested timing: 2 minutes on collective impact and key metrics, 9 minutes on three featured stories (about 3 each), and 4 minutes on additional accomplishments.

United Rentals: from experimenting with AI to building the foundations of an AI-enabled delivery lifecycle. 90%+ automation coverage for Cart regression by October; 320+ automated tests for Australia and New Zealand readiness; a visual regression capability covering 257 page scenarios across 8 device profiles (about 2,000 screenshots in under 40 minutes); AI-supported acceptance-criteria validation on pull requests with human approval gates; and Churchill AI-DLC expanding to more engineering teams. Why it matters: we're helping clients transform how they deliver software, not just speed up tasks.

Sherwin-Williams: turning Confidence standards into the operating model for enterprise transformation. Standardized Agile, Jira, requirements and delivery expectations adopted across six T4 product teams in August; three releases in six weeks (Design+ Dashboard, Pro+ Digital Account Opening, Color Expert Exterior Visualizer); AI-assisted requirements analysis generating executable Playwright automation inside the Definition of Done. Note: reports reference five pods, six product teams and seven workstreams at different points; reconcile before picking a headline metric.

Mazda: creating clarity and confidence across an interconnected enterprise program. AEM modernization milestones; 30 Vehicle Landing Pages migrated (October); program-level AQ governance across AEM, CRM, Incentives, BAU and Analytics; the Mazda Open Intelligence Hub and self-service Sidekick Agents; evolving Agent Ops governance. Note: validate the October 20 and 27 milestones before presenting them as complete.

Decision needed: are these the right three featured stories?
""")
stories = [
    ("United Rentals", "From AI experimentation to an AI-enabled delivery lifecycle", "90%+",
     "automation coverage for Cart regression (Oct)", [
         "320+ automated tests for AU/NZ readiness",
         "Visual regression: 257 scenarios, 8 devices, ~2,000 screenshots in <40 min",
         "Churchill AI-DLC expanding to more teams",
     ], None),
    ("Sherwin-Williams", "Confidence standards as the operating model for T4", "6",
     "T4 product teams on shared Agile standards (Aug)", [
         "3 releases in 6 weeks",
         "AI-assisted requirements to executable Playwright tests in the Definition of Done",
         "New ways of working for handoffs, estimation, dependencies",
     ], "pods vs. teams vs. workstreams"),
    ("Mazda", "Clarity and confidence across an interconnected program", "30",
     "Vehicle Landing Pages migrated (Oct)", [
         "Program-level AQ governance across AEM, CRM, Incentives, BAU, Analytics",
         "Open Intelligence Hub and Sidekick Agents",
         "Evolving Agent Ops governance",
     ], "Confirm Oct 20 and 27 milestones"),
]
cw = (W - 2 * 0.3) / 3
for i, (client, headline, stat, stat_lbl, items, check) in enumerate(stories):
    x = L + i * (cw + 0.3)
    text(s, x, TOP, cw, 0.4, client, size=22, font=HEAD)
    text(s, x, TOP + 0.45, cw - 0.1, 0.5, headline, size=13, color=MGRAY, line=1.0)
    rule(s, x, TOP + 1.0, x + cw, TOP + 1.0, color=BLACK)
    text(s, x, TOP + 1.1, 2.0, 0.8, stat, size=44, font=HEAD, color=BLACK)
    text(s, x + 2.0, TOP + 1.3, cw - 2.0, 0.6, stat_lbl, size=12, color=MGRAY, anchor=MSO_ANCHOR.MIDDLE, line=1.0)
    text(s, x, TOP + 2.1, cw - 0.1, 2.2, [{"t": it, "check": True} for it in items], size=13, after=6, line=1.05)
    if check:
        tag = box(s, x, TOP + 3.9, cw, 0.36, fill=LGOLD)
        label_in(tag, f"To validate: {check}", size=11, color=BLACK, font=BODY, align=PP_ALIGN.LEFT)
        tag.text_frame.margin_left = Inches(0.12)
b = box(s, L, 6.45, W, 0.55, fill=BLACK)
text(s, L + 0.25, 6.45, W - 0.5, 0.55, [{"t": [
    {"t": "Within 15 minutes:  ", "font": HEAD, "color": WHITE},
    {"t": "2 min collective impact and metrics   ·   9 min featured stories (~3 each)   ·   4 min additional accomplishments",
     "color": LGRAY},
]}], size=13, anchor=MSO_ANCHOR.MIDDLE)

# =======================================================================================
# 8 — OUR IMPACT: THEMES AND BREADTH
# =======================================================================================
s = content("Organize the impact report by theme, and show the breadth beyond our largest accounts", """
Left: the five impact themes come straight from the Technology Reports. Building the report around them makes it a practice story rather than a stack of team updates.

Right: the additional accomplishments make sure the report represents the whole practice, not only the biggest accounts. Not every item needs presentation time; a visual summary or rotating collage can show the breadth in a couple of minutes. GDMS, Conagra and Spire also contributed platform improvements, content initiatives, technical SEO, maintenance and quality work.

The working content target is three featured stories plus six to eight additional highlights.

Decision needed: are there other client or practice accomplishments that deserve more visibility?
""")
text(s, L, TOP, 5.2, 0.35, "Five impact themes", size=16, font=HEAD)
imp = [
    ("Delivering real outcomes", "Launches, modernized platforms, milestones"),
    ("Reinventing delivery", "AI-assisted workflows and intelligent automation"),
    ("Building confidence at scale", "Quality engineering, governance, validation"),
    ("Establishing repeatable practices", "Shared Agile standards, documentation, onboarding"),
    ("Growing as strategic partners", "Client trust and expanded responsibilities"),
]
for i, (h, d) in enumerate(imp):
    y = TOP + 0.5 + i * 0.88
    text(s, L, y, 0.55, 0.4, f"{i + 1:02d}", size=16, color=DGOLD, font=HEAD)
    text(s, L + 0.55, y, 4.6, 0.35, h, size=15, font=HEAD)
    text(s, L + 0.55, y + 0.36, 4.6, 0.4, d, size=13, color=MGRAY)
    rule(s, L, y + 0.8, L + 5.2, y + 0.8)
x2 = 5.95
text(s, x2, TOP, R - x2, 0.35, "Additional highlights to choose from", size=16, font=HEAD)
extra = [
    ("Pokémon Automated Retail", "Pilot soft-launch across 15 vending machines"),
    ("Tennessee Tourism", "16,000+ initial requests for the 2026 Vacation Guide"),
    ("Coca-Cola", "America 250 activation on an 8-week timeline"),
    ("US Navy", "Clearer priorities and visibility across six workstreams"),
    ("EchoPark", "First platform release supporting 11 customer journeys"),
    ("Mayo Clinic", "Expanded technical role; first production content migrations"),
    ("Thrivent", "Investor experience improvements, incl. ETF Performance Page"),
    ("Global Payments", "Enterprise CMS recommendation after acquisition"),
    ("Boehringer Ingelheim", "Modernized AQ tooling and process with Jira and Zephyr"),
    ("GDMS, Conagra, Spire", "Platform, content, technical SEO and quality work"),
]
tw, th, g = (R - x2 - 0.2) / 2, 0.8, 0.1
for i, (c, d) in enumerate(extra):
    col, row = i % 2, i // 2
    x = x2 + col * (tw + 0.2)
    y = TOP + 0.5 + row * (th + g)
    box(s, x, y, tw, th, line=LGRAY)
    text(s, x + 0.15, y + 0.1, tw - 0.3, 0.3, c, size=13, font=HEAD)
    text(s, x + 0.15, y + 0.36, tw - 0.3, 0.35, d, size=12, color=MGRAY, line=1.0)

# =======================================================================================
# 9 — OUR PEOPLE
# =======================================================================================
s = content("Our People (8 min): short, genuine spotlights tied to real contributions", """
Keep the recognition and shout-outs that are traditionally part of the year-end All Hands, but connect them more directly to meaningful contributions.

Recognition themes could include innovation and problem-solving, leadership through difficult delivery, cross-team collaboration, mentorship and knowledge sharing, improving our craft and ways of working, and helping clients and colleagues succeed.

The examples on the right come from the Technology Reports. They are starting points identified through reporting, not finalized nominations or selections.

Preparation: ask each group or pod lead to nominate individuals or teams with a short explanation of what they accomplished, why it mattered and who benefited. Avoid a competitive awards format; a series of short spotlights will feel more authentic.

Decision needed: how should we collect and select nominations so representation is broad, including less visible contributors?
""")
text(s, L, TOP, 4.4, 0.35, "How we'll approach it", size=16, font=HEAD)
text(s, L, TOP + 0.5, 4.4, 4.2, [
    {"t": "Each group and pod lead nominates individuals or teams: what they did, why it mattered, who benefited",
     "check": True},
    {"t": "A series of short spotlights, not a competitive awards program", "check": True},
    {"t": "Recognize growth, collaboration and resilience, not just output", "check": True},
    {"t": "Ask leads directly so less visible contributors aren't missed", "check": True},
], size=14, after=10, line=1.05)
x2 = 5.1
text(s, x2, TOP, R - x2, 0.35, "Starting points from the Technology Reports", size=16, font=HEAD)
table(s, x2, TOP + 0.45, R - x2, [2.55, R - x2 - 2.55], [
    ("Theme", "Example contributors"),
    ("AI-enabled delivery", "Kunal Pacheriwala: expanding the Churchill AI development platform"),
    ("Quality innovation", "Jacob Barrett: visual regression architecture; United Rentals automation team"),
    ("Program coordination", "Pan Leung: Navy workstream visibility and prioritization"),
    ("Complex delivery leadership", "Melissa Jones: Mazda cross-squad coordination and risk mitigation"),
    ("Practice standardization", "Sherwin-Williams T4 analyst team, incl. Abbey Logan on dependency management"),
    ("Knowledge sharing", "Spencer Brown: AI agent workshop and follow-up coaching"),
    ("Quality leadership", "Vanessa Lacy: Mazda program-level AQ; Kristy Riley: process modernization"),
    ("Community engagement", "Jessica Ward, Sharee Thompson and volunteers: KC Tech Council LEGiT program"),
], row_h=0.45, size=12)
text(s, x2, 6.68, R - x2, 0.35, "Examples identified through reporting, not finalized nominations.", size=12, color=MGRAY)

# =======================================================================================
# 10 — OUR FUTURE: THE CONFIDENCE FUTURE 3
# =======================================================================================
s = content("Looking ahead (5 min): the Confidence Future 3", """
Use VML Intelligence's Future 100: 2026 as inspiration for three original predictions about the future of Confidence. The goal isn't to recap the report or claim certainty about 2027; it's to introduce possibilities that challenge how we think about our work.

Future 1, AI as a team member. Inspiration: Synthetic Generation (pp. 50–52) and Democratizing creAItivity (pp. 100–101). 2026 connection: Churchill AI-DLC, AI-assisted requirements analysis, quality lifecycle agents, Agent Ops governance. Our roles may evolve toward orchestrating human and AI contributors.

Future 2, trust as our deliverable. Inspiration: Truth Literacy (pp. 23–25), Omnisurveillance (pp. 53–55), Coded Empathy (pp. 271–272). 2026 connection: continuous quality engineering, AI-assisted validation, program-level AQ governance, privacy and agent approval processes. Clients may need assurance that autonomous systems are reliable, explainable and well governed.

Future 3, beyond the website. Inspiration: The Reality Interface (pp. 61–63), Acommerce (pp. 196–198), Digital Intent (pp. 56–58). 2026 connection: conversational AI, natural-language search, agent-enabled experiences. Conventional web journeys may become one of several ways customers interact with brands.

Alternative theme: what if we reinvented Confidence from scratch? This open question may suit the popcorn conversation especially well.

Presentation approach: borrow the Future 100's editorial format (a bold prediction, a short trend explanation, one example or data point, a connection to our 2026 work, and one provocative question), with strong visuals rather than dense technical slides.

Decision needed: do these three predictions represent the right mix?
""", dark=True)
fut = [
    ("What if AI becomes a member of every delivery team?",
     "Churchill AI-DLC, AI-assisted requirements, quality lifecycle agents, Agent Ops",
     "If we added an AI agent to a Scrum team tomorrow, what would we give it, and what would we never delegate?"),
    ("What if trust becomes our most valuable deliverable?",
     "Continuous quality engineering, AI-assisted validation, program-level AQ governance",
     "What would Confidence look like if clients hired us primarily to establish trust in their technology?"),
    ("What if the website is no longer the destination?",
     "Conversational AI, natural-language search, agent-enabled experiences",
     "How do we define, design and validate experiences that don't follow a predictable journey?"),
]
cw = (W - 2 * 0.4) / 3
for i, (pred, conn, q) in enumerate(fut):
    x = L + i * (cw + 0.4)
    numeral(s, x, TOP, i + 1, color=GOLD, size=40)
    text(s, x, TOP + 0.8, cw - 0.1, 1.3, pred, size=22, color=WHITE, font=HEAD, line=1.05)
    rule(s, x, 4.2, x + cw, 4.2, color=MGRAY)
    text(s, x, 4.32, cw, 0.3, "2026 CONNECTION", size=10, color=GRAY, font=HEAD)
    text(s, x, 4.6, cw - 0.1, 0.65, conn, size=12, color=LGRAY, line=1.05)
    text(s, x, 5.3, cw - 0.1, 1.0, [{"t": f"“{q}”", "italic": True}], size=13, color=LGOLD, line=1.1)
text(s, L, 6.55, W, 0.45, [{"t": [
    {"t": "Alternative:  ", "font": HEAD, "color": WHITE},
    {"t": "what if we reinvented Confidence from scratch?   ", "color": LGRAY},
    {"t": "Inspired by VML Intelligence, The Future 100: 2026.", "color": GRAY},
]}], size=12, anchor=MSO_ANCHOR.MIDDLE)

# =======================================================================================
# 11 — POPCORN CONVERSATION
# =======================================================================================
s = content("Popcorn conversation (12 min): the part most likely to make this year feel different", """
An informal, imaginative discussion where representatives from across Confidence react to an idea and build on each other's thoughts: closer to spontaneous brainstorming than a panel.

Recommended format: four or five people representing different disciplines, pods and delivery experiences. A moderator introduces a provocative scenario and invites people to build on each other. For example, an analyst imagines requirements for an autonomous digital experience; a quality engineer builds on that by describing how it could be validated; a Scrum Master explores how the team would adapt; someone else considers new client opportunities.

Facilitation guidelines: keep it exploratory, not evaluative. Encourage "yes, and." Make room for different disciplines. Don't dismiss ideas because of current constraints. No consensus, business case or recommendations required. Capture interesting possibilities without creating obligations.

Alternative: two or three cross-practice small groups, each on a different scenario, sharing their most interesting idea back. Broader participation, but it needs more time and facilitation, so extend the meeting to about 70 minutes if we choose it.

Strong recommendation: protect this segment from being cut as other presenters ask for time. It's what will make this All Hands feel different, gives new voices a chance to contribute, and brings Kylie's vision to life.

Decisions needed: one shared scenario or several questions? Live conversation, breakouts or a combination? How much prep do participants get? Who captures the ideas?
""")
text(s, L, TOP, 6.6, 0.35, "Recommended format: a live, lightly facilitated conversation", size=16, font=HEAD)
text(s, L, TOP + 0.5, 6.6, 2.2, [
    {"t": "4–5 voices across disciplines, pods and delivery experiences", "check": True},
    {"t": "A moderator sets up one provocative scenario, then invites people to build on each other", "check": True},
    {"t": "“Yes, and”: exploratory, not evaluative; no consensus or business case required", "check": True},
    {"t": "Audience joins through chat; one takeaway captured at the end", "check": True},
], size=14, after=8, line=1.05)
plan = [(2, "Set up the scenario"), (7, "Popcorn conversation"), (2, "Audience via chat"), (1, "Capture")]
unit = 6.6 / 12
x, by = L, 5.1
text(s, L, by - 0.45, 6.6, 0.3, "12-minute structure", size=12, color=MGRAY, font=HEAD)
for mins, lbl in plan:
    seg = box(s, x, by, mins * unit - 0.04, 0.5, fill=GOLD if mins == 7 else BLACK)
    label_in(seg, f"{mins}′", size=12, color=BLACK if mins == 7 else WHITE)
    text(s, x, by + 0.58, max(mins * unit - 0.08, 1.0), 0.5, lbl, size=11, color=MGRAY, line=1.0)
    x += mins * unit
text(s, L, 6.5, 6.6, 0.5, [{"t": [
    {"t": "Alternative: ", "font": HEAD},
    {"t": "2–3 small-group breakouts. Broader participation, but plan for ~70 minutes."},
]}], size=12, color=BLACK, line=1.05)
x2 = 7.35
b = box(s, x2, TOP, R - x2, 2.3, fill=BLACK)
text(s, x2 + 0.3, TOP + 0.25, R - x2 - 0.6, 1.9, [
    {"t": "Recommendation: protect these 12 minutes", "font": HEAD, "size": 18, "color": LGOLD, "after": 8},
    {"t": "When other presenters ask for time, take it from elsewhere. This is what will make this All Hands feel "
          "different, give new voices a chance to contribute, and bring Kylie's vision to life.",
     "size": 13, "color": WHITE},
], line=1.1)
text(s, x2, 4.5, R - x2, 0.35, "Prompts to choose from", size=14, font=HEAD)
text(s, x2, 4.9, R - x2, 2.1, [
    {"t": "If we were building Confidence from scratch today, what would we do differently?", "check": True},
    {"t": "What parts of our delivery model might become unnecessary?", "check": True},
    {"t": "Where could Technology Analysis, Scrum and Quality Engineering converge?", "check": True},
    {"t": "What would we pursue without today's roles, tools or processes?", "check": True},
], size=13, after=6, line=1.05)

# =======================================================================================
# 12 — ROLES
# =======================================================================================
s = content("Roles: Travis leads planning, with a small working group behind him", """
Planning ownership: Travis Kubicek is proposed as planning lead (to confirm), with Zack Jemas supporting on agenda development, content planning and coordination.

Important: we should set up a small planning group rather than leaving all of the presentation production to Travis. Group and pod leads are the main source of impact examples, metric validation and nominations.

Hosting rotation: Kylie would like to give someone new from AQ the chance to host, rather than relying on Nick Dalton again. Names mentioned so far include Dianna, Christie, Vanessa and Jessica; the final selection is open.

Decisions needed: Can Travis serve as planning lead with a small working group? Who should represent AQ as host or co-host? Which team members would be good contributors to the impact and future-focused segments?
""")
table(s, L, TOP, 8.3, [2.75, 5.55], [
    ("Role", "Responsibility"),
    ("Kylie Schleicher", "Overall direction, leadership alignment and feedback"),
    ("Travis Kubicek (to confirm)", "Planning lead: meeting planning and coordination"),
    ("Zack Jemas", "Agenda development, content planning and coordination support"),
    ("Group and pod leads", "Impact examples, metric validation, nominations, participant ideas"),
    ("Confidence leadership", "Leadership updates"),
    ("AQ host (new voice)", "Co-hosting or facilitation"),
    ("Future Lab working group", "Refine future themes, discussion format and participants"),
    ("Presentation contributors", "Content consolidation, design, technical coordination, rehearsal"),
], row_h=0.53, size=13, highlight=2)
x2 = 9.0
box(s, x2, TOP, R - x2, 2.25, line=BLACK)
text(s, x2 + 0.25, TOP + 0.22, R - x2 - 0.5, 1.9, [
    {"t": "Don't put it all on one person", "font": HEAD, "size": 16, "after": 8},
    {"t": "A small planning group shares content consolidation, design and rehearsal, so production doesn't sit "
          "with Travis alone.", "size": 13},
], line=1.1)
box(s, x2, TOP + 2.5, R - x2, 2.25, fill=LGOLD)
text(s, x2 + 0.25, TOP + 2.72, R - x2 - 0.5, 1.9, [
    {"t": "A new host from AQ", "font": HEAD, "size": 16, "after": 8},
    {"t": "Kylie would like to give someone new the opportunity, rather than Nick Dalton again. Mentioned so far: "
          "Dianna, Christie, Vanessa, Jessica. Selection is open.", "size": 13},
], line=1.1)

# =======================================================================================
# 13 — TIMELINE
# =======================================================================================
s = content("Timeline: five weeks from alignment to the All Hands", """
Walk the timeline left to right. The key dependency is November reporting: Kylie's preference is to hold the meeting shortly after November reporting, ideally before Thanksgiving, and before the end-of-year rush.

If Nov 23–24 isn't practical, the preferred fallback is the first week of December. Attendance, leadership availability and contributor preparation will drive the final date.

The first two steps are this week and next: leads review and alignment (today), then planning kickoff with the planning team, host, contributors and the content-collection approach.

Decision needed: can we confirm a target date before Thanksgiving?
""")
steps = [
    ("Oct 19–23", "Leads review and alignment", "Approved direction, roles, target date"),
    ("Oct 26–30", "Planning kickoff", "Planning team, host, contributors, collection plan"),
    ("Nov 2–6", "Initial content collection", "Impact examples, metrics, nominations, future themes"),
    ("Nov 9–13", "Presentation development", "Draft impact report, recognition, Future 3 slides"),
    ("Nov 16–18", "November reporting integration", "Validated year-end highlights and data"),
    ("Nov 19–20", "Final preparation", "Review, facilitation walkthrough, tech rehearsal"),
    ("Nov 23–24", "Year-End All Hands", "Run the meeting"),
    ("After", "Follow-up", "Recap, recognize contributors, capture future ideas"),
]
n = len(steps)
cw = W / n
ly = TOP + 0.75
rule(s, L, ly, R, ly, color=BLACK, width=1)
for i, (when, what, out) in enumerate(steps):
    x = L + i * cw
    key = when == "Nov 23–24"
    now = i == 0
    text(s, x, TOP, cw - 0.15, 0.5, when, size=16, font=HEAD, color=DGOLD if key else BLACK)
    box(s, x, ly - 0.11, 0.22, 0.22, fill=GOLD if key else (BLACK if now else WHITE), line=BLACK)
    text(s, x, ly + 0.35, cw - 0.2, 0.9, what, size=14, font=HEAD, line=1.0)
    text(s, x, ly + 1.3, cw - 0.2, 1.4, out, size=12, color=MGRAY, line=1.05)
text(s, L + 0.32, ly - 0.42, 1.0, 0.25, "We are here", size=10, color=MGRAY)
box(s, L, 5.75, W, 1.25, line=LGRAY)
text(s, L + 0.3, 5.9, W - 0.6, 1.0, [
    {"t": "Timing considerations", "font": HEAD, "size": 14, "after": 4},
    {"t": "Kylie prefers to meet shortly after November reporting and before Thanksgiving, ahead of the end-of-year rush. "
          "Fallback: first week of December. Attendance, leadership availability and contributor prep will set the final date.",
     "size": 13, "color": MGRAY},
], line=1.05)

# =======================================================================================
# 14 — WHAT WE NEED FROM LEADS
# =======================================================================================
s = content("What we need from group and pod leads: a light ask, not another report", """
Preparation should lean on the existing Technology Reports and the November reporting cycle wherever possible. The intent is a light, repeatable collection process, not another full reporting exercise.

The six inputs on the left are what we'll ask each group and pod lead for during the Nov 2–6 collection window.

Rules for metrics: only use measures that can be reasonably validated and explained. The reports span multiple cycles and often reference the same engagements, so recurring accomplishments should be deduplicated rather than added together. Keep completed outcomes separate from planned launches, work in progress and forecasted benefits. Some engagements, including USMC, report in separate sources, so the five Technology Reports are a strong starting point rather than a complete annual inventory.

Potential practice-wide measures: number of clients or accounts supported; product releases, launches or major milestones; significant automation coverage and efficiency gains; new or expanded responsibilities in client engagements; reusable frameworks and accelerators; practice development and learning activities.

Decisions needed: which practice-wide metrics can we reliably obtain, and how do we validate existing report data against November reporting?
""")
text(s, L, TOP, 7.3, 0.35, "Six inputs per group or pod  ·  due Nov 2–6", size=16, font=HEAD)
inputs = [
    ("Top accomplishments", "Two or three meaningful achievements"),
    ("Evidence of impact", "Quantitative results, client feedback, observable outcomes"),
    ("Client contribution", "What changed or improved because of our involvement"),
    ("Practice evolution", "New processes, skills, technology or ways of working"),
    ("Recognition", "Individuals or teams who made a meaningful contribution"),
    ("Future opportunities", "Emerging ideas or capabilities worth exploring"),
]
tw, th = (7.3 - 0.25) / 2, 1.3
for i, (h, d) in enumerate(inputs):
    col, row = i % 2, i // 2
    x = L + col * (tw + 0.25)
    y = TOP + 0.5 + row * (th + 0.2)
    box(s, x, y, tw, th, line=BLACK)
    text(s, x + 0.2, y + 0.15, 0.6, 0.4, f"{i + 1:02d}", size=14, color=DGOLD, font=HEAD)
    text(s, x + 0.2, y + 0.5, tw - 0.4, 0.35, h, size=15, font=HEAD)
    text(s, x + 0.2, y + 0.85, tw - 0.4, 0.45, d, size=12, color=MGRAY, line=1.0)
x2 = 8.05
box(s, x2, TOP, R - x2, 4.8, fill=BLACK)
text(s, x2 + 0.3, TOP + 0.25, R - x2 - 0.6, 4.7, [
    {"t": "Rules for metrics", "font": HEAD, "size": 16, "color": LGOLD, "after": 10},
    {"t": "Only use what can be validated and explained", "check": True},
    {"t": "Deduplicate wins repeated across reporting cycles", "check": True},
    {"t": "Separate completed outcomes from planned, in-progress and forecast work", "check": True},
    {"t": "Reconcile scope before choosing a headline number", "check": True},
    {"t": "Fill reporting gaps (e.g. USMC reports separately) with direct lead input", "check": True},
], size=13, color=WHITE, after=8, line=1.05)

# =======================================================================================
# 15 — RISKS
# =======================================================================================
s = content("Main risks are tone, balance and time, and each one has a clear mitigation", """
These are the risks most likely to affect how the meeting lands. Two are worth emphasizing.

Tone: an overly celebratory meeting could feel disconnected from the year people actually had. Ground recognition in real accomplishments and acknowledge the challenges openly.

Time: as presenters ask for more time, the future conversation is usually the first thing cut. Protect it.

Other considerations: if small-group breakouts are chosen, logistics could consume the 60 minutes, so use a live conversation for 60 or extend to about 70 minutes. Reporting gaps could distort the story, so verify scope and get supplemental input where major teams or accounts are missing.
""")
table(s, L, TOP, W, [4.6, W - 4.6], [
    ("Risk", "Mitigation"),
    ("Meeting feels overly celebratory", "Ground recognition in real accomplishments and acknowledge the year's challenges"),
    ("Impact report becomes a status update", "Focus on meaningful outcomes and a few illustrative stories"),
    ("Reporting favors the largest accounts", "Include additional highlights and nominations from across the practice"),
    ("Data is inconsistent or unavailable", "Use validated, well-defined metrics, with qualitative evidence where needed"),
    ("AI content overshadows other work", "Balance innovation with delivery leadership, analysis, quality and collaboration"),
    ("Future discussion feels generic", "Tie provocations to real 2026 experience; ask unconventional questions"),
    ("Popcorn conversation is awkward or unfocused", "An engaging moderator, varied participants, a few prepared prompts"),
    ("Preparation becomes burdensome", "Reuse existing reporting and spread content responsibilities"),
    ("Future ideas don't go anywhere", "Document promising concepts and name an owner for optional follow-up"),
], row_h=0.5, size=13)

# =======================================================================================
# 16 — DECISIONS
# =======================================================================================
s = content("Decisions for today", """
The primary objective of this meeting is alignment on the concept, not a detailed content review.

The six decisions that let planning start:
1. Approve the overall Impact, People and Future structure.
2. Confirm the target date and meeting duration.
3. Confirm the planning lead and the small supporting team.
4. Identify an AQ host and initial contributors.
5. Agree on how we collect impact data and recognition nominations.
6. Authorize a small cross-practice group to refine the future themes and conversation format.

Once aligned, the planning team can begin consolidating impact stories, confirming metrics, collecting nominations and developing the Future Lab content, starting with kickoff the week of Oct 26.

The full list of open questions is in the appendix.
""", dark=True)
dec = [
    ("Approve the structure", "Impact, People and Future, with 10 min for leadership updates"),
    ("Confirm date and duration", "Nov 23 or 24 (fallback: first week of Dec); 60 or ~70 minutes"),
    ("Confirm the planning lead", "Travis, plus a small supporting team"),
    ("Identify an AQ host", "And initial contributors for the impact and future segments"),
    ("Agree the collection approach", "Impact data, metric validation and recognition nominations"),
    ("Authorize a Future Lab group", "A small cross-practice group refines the future themes and format"),
]
cw, rh = (W - 0.5) / 2, 1.3
for i, (h, d) in enumerate(dec):
    col, row = i // 3, i % 3
    x = L + col * (cw + 0.5)
    y = TOP + row * rh
    text(s, x, y, 0.8, 0.6, f"{i + 1:02d}", size=28, color=GOLD, font=HEAD)
    text(s, x + 0.95, y + 0.02, cw - 0.95, 0.4, h, size=20, color=WHITE, font=HEAD)
    text(s, x + 0.95, y + 0.48, cw - 0.95, 0.5, d, size=13, color=LGRAY)
    rule(s, x, y + rh - 0.15, x + cw, y + rh - 0.15, color=MGRAY)
text(s, L, 6.1, W, 0.8, [{"t": [
    {"t": "Next:  ", "font": HEAD, "color": LGOLD},
    {"t": "planning kickoff Oct 26–30, then content collection Nov 2–6. Full list of open questions in the appendix.",
     "color": WHITE},
]}], size=14)

# =======================================================================================
# 17 — CLOSING STATEMENT
# =======================================================================================
s = new_slide(DARK, "Divider - Statement (Dark)", keep=(0,), notes="""
Close on the intent behind the whole plan.

This is a chance to do something more meaningful than a conventional year-end celebration or a review of project accomplishments: to show how much Confidence contributes across our clients, acknowledge the people who make that work possible, and start a broader conversation about the future of our practice.

The 2026 Technology Reports give us credible evidence that we're already evolving. AI-assisted delivery, more mature quality engineering, standardized ways of working and expanded client partnerships aren't abstract ambitions; they're coming out of work our teams are doing today. The Future 100 gives us a way to look beyond that and imagine what comes next.
""")
t = ph(s, 0)
t.left, t.top, t.width, t.height = Inches(L), Inches(1.0), Inches(11.8), Inches(3.6)
t.text_frame.text = ("Finish the year with appreciation for what we've accomplished, pride in how we've grown, "
                     "and genuine curiosity about what we could become.")
for p in t.text_frame.paragraphs:
    for r in p.runs:
        r.font.size = Pt(48)
text(s, L, 5.3, 11.0, 0.5, [{"t": [
    {"t": "Recognize our impact.  Appreciate our people.  ", "color": LGRAY},
    {"t": "Imagine our future.", "color": GOLD},
]}], size=20)

# =======================================================================================
# APPENDIX
# =======================================================================================
APPENDIX_START = len(prs.slides._sldIdLst)
s = new_slide(DARK, "Divider (Dark)", title="Appendix", notes="""
Reference material for the planning team. Not intended to be presented.
""")

# A1 — Objectives, content targets, success ------------------------------------------------
s = content("Appendix: objectives, content targets and measures of success", """
Reference for the planning team: what the meeting should accomplish, the working content target, and how we'll know it worked.

The meeting is not intended to set a formal 2027 strategy or create new delivery commitments.

The objective is not to maximize the number of metrics, slides, recognitions or ideas. The strength of the meeting will come from the clarity, authenticity and relevance of the experience.
""")
cols = [
    ("Meeting objectives", [
        "Make our impact visible, with credible measures and examples",
        "Recognize meaningful contributions genuinely and inclusively",
        "Connect work across pods, clients and disciplines into one practice story",
        "Inspire forward-looking thinking that challenges assumptions",
        "Generate a few ideas worth exploring in 2027",
    ], "Not a formal 2027 strategy or new delivery commitments."),
    ("Working content target", [
        "Three featured impact stories",
        "Six to eight additional accomplishment highlights",
        "A representative group of contributor shout-outs",
        "Three Future 100-inspired provocations",
        "One collaborative popcorn-style conversation",
        "A brief recap of ideas to carry into 2027",
    ], None),
    ("We'll call it a success if it", [
        "Shows a compelling, credible picture of our 2026 impact",
        "Recognizes meaningful individual and team contributions",
        "Shows how our capabilities and ways of working evolved",
        "Shares important context from leadership",
        "Creates connection across disciplines and pods",
        "Sparks imaginative discussion and a few ideas to revisit",
        "Leaves people feeling appreciated and connected",
    ], None),
]
cw = (W - 2 * 0.35) / 3
for i, (head, items, foot) in enumerate(cols):
    x = L + i * (cw + 0.35)
    text(s, x, TOP, cw, 0.35, head, size=16, font=HEAD)
    rule(s, x, TOP + 0.48, x + cw, TOP + 0.48, color=BLACK)
    text(s, x, TOP + 0.65, cw, 4.0, [{"t": it, "check": True} for it in items], size=13, after=6, line=1.05)
    if foot:
        text(s, x, 6.5, cw, 0.5, foot, size=12, color=MGRAY)

# A2 — Open questions ----------------------------------------------------------------------
s = content("Appendix: open questions for leads", """
The full set of open questions from the planning proposal, grouped by topic. Today's meeting should focus on the direction and ownership questions; the rest can be worked by the planning team and the Future Lab working group.
""")
qs = [
    ("Direction and structure", [
        "Do we agree with the Impact, People and Future narrative?",
        "Is the balance of updates, reporting, recognition and discussion right?",
        "Keep 60 minutes, or extend for more audience participation?",
        "Can we confirm a target date before Thanksgiving?",
    ]),
    ("Leadership and ownership", [
        "Who presents leadership updates, and on which topics?",
        "Can Travis serve as planning lead, with a small working group?",
        "Who should represent AQ as host or co-host?",
        "Who would be good contributors to the impact and future segments?",
    ]),
    ("Impact and recognition", [
        "Which practice-wide metrics can we reliably obtain?",
        "Are United Rentals, Sherwin-Williams and Mazda the right three stories?",
        "Which other accomplishments deserve more visibility?",
        "How do we collect nominations for broad representation?",
        "How do we validate reports against November reporting?",
    ]),
    ("Future discussion", [
        "Are the three predictions the right mix of topics?",
        "One shared scenario, or multiple questions?",
        "Live popcorn, breakouts or a combination?",
        "How much preparation do participants get?",
        "Who captures the ideas and decides what's worth exploring?",
    ]),
]
cw = (W - 3 * 0.3) / 4
qn_ = 1
for i, (head, items) in enumerate(qs):
    x = L + i * (cw + 0.3)
    text(s, x, TOP, cw, 0.35, head, size=15, font=HEAD)
    rule(s, x, TOP + 0.48, x + cw, TOP + 0.48, color=BLACK)
    paras = []
    for it in items:
        paras.append({"t": [{"t": f"{qn_:02d}  ", "font": HEAD, "color": DGOLD}, {"t": it}], "after": 9})
        qn_ += 1
    text(s, x, TOP + 0.65, cw, 4.4, paras, size=13, line=1.05)

# A3 — Sources -------------------------------------------------------------------------------
s = content("Appendix: source material", """
All final presentation materials should be validated with the respective contributors and leads before publication.
""")
src = [
    ("Leadership discussion", ["Zack and Kylie's All Hands planning conversation, October 2026"]),
    ("Internal reporting", [
        "Technology Reports: February, April, June, August and October 2026",
        "November 2026 report, when available",
        "Primary review focus: Jeff Mullennix, Kirby Hamilton, Zack Jemas, Travis Kubicek, Hilary Heugh, "
        "Hannah Anderson, Ann Millington, Dianna Keen, Irina Smith, Kylie Schleicher, Nick Dalton, Vanessa Lacy, "
        "Andie Johnson and Beth Noser",
    ]),
    ("Future inspiration", [
        "VML Intelligence, The Future 100: 2026",
        "Sections: Truth Literacy, Synthetic Generation, Digital Intent, The Reality Interface, Democratizing creAItivity, "
        "Acommerce, Omnisurveillance, Coded Empathy",
    ]),
]
cw = (W - 2 * 0.35) / 3
for i, (head, items) in enumerate(src):
    x = L + i * (cw + 0.35)
    text(s, x, TOP, cw, 0.35, head, size=16, font=HEAD)
    rule(s, x, TOP + 0.48, x + cw, TOP + 0.48, color=BLACK)
    text(s, x, TOP + 0.65, cw, 3.8, [{"t": it, "check": True} for it in items], size=13, after=8, line=1.05)
text(s, L, 6.5, W, 0.5, "Validate all final materials with the respective contributors and leads before publication.",
     size=12, color=MGRAY)

# --- sections for easy navigation in PowerPoint --------------------------------------------
SECTIONS = [("Proposal", 0), ("The meeting", 4), ("Planning", 11), ("Decisions", 15), ("Appendix", APPENDIX_START)]
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
bounds = [i for _, i in SECTIONS] + [len(ids)]
for k, (name, start) in enumerate(SECTIONS):
    sec = etree.SubElement(secLst, f"{{{P14}}}section", name=name, id="{" + str(uuid.uuid4()).upper() + "}")
    lst = etree.SubElement(sec, f"{{{P14}}}sldIdLst")
    for sid in ids[start:bounds[k + 1]]:
        etree.SubElement(lst, f"{{{P14}}}sldId", id=sid)

prs.core_properties.title = "Confidence 2026: Our Impact, Our Future"
prs.core_properties.subject = "Year-End All Hands planning proposal for Confidence Leads"
prs.core_properties.author = "Zack Jemas"
prs.save(str(OUT))
print(f"Saved {OUT}  ({len(ids)} slides)")
