"""Build the Confidence 2026 Year-End All Hands content outline on the VML 2026 template.

Usage: python3 build_deck.py [output.pptx]

Every slide sits on a native VML layout (light and dark masters, header rules, spark mark,
slide number), and all type uses the template's Snowflake Sans theme fonts. Color follows the
brand rule: black and white first, with one supporting hue (Gold) used sparingly.
"""
import copy
import sys
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
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE.parent / "Confidence-2026-All-Hands-Content-Outline.pptx"

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
# 1 — THE IDEA AND INTENTIONS
# =======================================================================================
s = content("What Kylie has in mind for our year-end All Hands", """
For our final Confidence All Hands of 2026, Kylie would like a slightly different approach from the traditional year-end celebration: a 2026 Impact Report and a 2027 Future Outlook.

The intentions behind it:
• There's plenty worth recognizing, but an overly celebratory tone could feel disconnected from the challenges, uncertainty and demands people experienced this year. Lead with the tangible contributions Confidence made.
• Make room for the reality of people's experiences while still creating something productive and forward-looking.
• Use VML Intelligence's Future 100 as inspiration to think beyond incremental process improvements and imagine how we might reinvent the practice.

Format: about 60 minutes, with roughly 10 minutes kept for Confidence leadership updates. Timing preference is shortly after November reporting and before Thanksgiving.

Planning details will live on a separate page; this deck is only the content outline and intent.
""", dark=True)
text(s, L, TOP, W, 0.8, "Confidence 2026: Our Impact, Our Future", size=44, color=WHITE, font=HEAD)
text(s, L, TOP + 0.9, W, 0.5, [{"t": [
    {"t": "Recognize our impact.  Appreciate our people.  ", "color": LGRAY},
    {"t": "Imagine our future.", "color": GOLD},
]}], size=22)
intents = [
    ("An impact report, not a party",
     "There's plenty to recognize, but an overly celebratory tone could feel out of step with the year. "
     "Lead with what Confidence actually contributed."),
    ("Honest about the year",
     "Make room for the reality of people's experiences, while still creating something productive and "
     "forward-looking."),
    ("Imagine, don't just improve",
     "Use VML's Future 100 to think beyond incremental process changes and imagine how we might reinvent "
     "the practice."),
]
cw = (W - 2 * 0.4) / 3
for i, (h, d) in enumerate(intents):
    x = L + i * (cw + 0.4)
    rule(s, x, 3.75, x + cw, 3.75, color=MGRAY)
    text(s, x, 3.95, 0.8, 0.5, f"{i + 1:02d}", size=16, color=GOLD, font=HEAD)
    text(s, x, 4.4, cw - 0.1, 0.5, h, size=20, color=WHITE, font=HEAD)
    text(s, x, 4.95, cw - 0.1, 1.3, d, size=14, color=LGRAY, line=1.15)
text(s, L, 6.55, W, 0.4, [{"t": [
    {"t": "Format:  ", "font": HEAD, "color": WHITE},
    {"t": "about 60 minutes, including ~10 minutes of Confidence leadership updates. "
          "Late November, shortly after November reporting.", "color": LGRAY},
]}], size=14)

# =======================================================================================
# 2 — THE OUTLINE
# =======================================================================================
s = content("The outline: what we did, who made it happen, and what could come next", """
A proposed flow for the hour. Timing is a starting point and can flex once leadership and presenter needs are known.

The arc matters more than the minutes: open honestly, give leadership a dedicated slot, show our collective impact, recognize the people behind it, then turn toward the future with a few provocations and a live conversation.

The bridge between the two halves: 2026 was a year of evolving from individual project successes toward reusable capabilities. If our work, capabilities and roles have already evolved this much, what could Confidence look like in 2027?

Suggestion: protect the popcorn conversation if other segments ask for more time. It's the part most likely to make this year's All Hands feel different and give new voices a chance to contribute.
""")
segs = [
    (5, "Welcome and opening", "Acknowledge the year; explain why this meeting looks different", "frame"),
    (10, "Leadership updates", "Two or three concise updates on the business, clients and practice", "frame"),
    (15, "Our Impact: 2026 in review", "A credible picture of the breadth, scale and value of our work", "core"),
    (8, "Our People: behind the impact", "Genuine recognition tied to real contributions", "core"),
    (5, "Looking ahead: the Confidence Future 3", "Three Future 100-inspired provocations", "core"),
    (12, "Popcorn conversation", "A live “yes, and” discussion about what Confidence could become", "key"),
    (5, "Closing", "Appreciation, a few takeaways, and how ideas will be captured", "frame"),
]
fills = {"frame": (LGRAY, BLACK), "core": (BLACK, WHITE), "key": (GOLD, BLACK)}
unit = W / 60
x = L
bar_y, bar_h = TOP + 0.05, 0.55
for mins, name, intent, kind in segs:
    f, c = fills[kind]
    seg = box(s, x, bar_y, mins * unit - 0.04, bar_h, fill=f)
    label_in(seg, f"{mins} min", size=12, color=c)
    x += mins * unit
ty = 3.05
rows_h = 0.56
text(s, L + 0.45, ty - 0.4, 1.3, 0.3, "Time", size=11, color=MGRAY, font=HEAD)
text(s, L + 1.8, ty - 0.4, 4.0, 0.3, "Segment", size=11, color=MGRAY, font=HEAD)
text(s, L + 6.3, ty - 0.4, 6.0, 0.3, "Intention", size=11, color=MGRAY, font=HEAD)
t0 = 0
for i, (mins, name, intent, kind) in enumerate(segs):
    y = ty + i * rows_h
    f, _ = fills[kind]
    box(s, L, y + 0.17, 0.2, 0.2, fill=f, line=BLACK if kind == "frame" else None)
    text(s, L + 0.45, y + 0.12, 1.3, 0.35, f"{t0:02d}–{t0 + mins:02d}", size=14, color=MGRAY)
    text(s, L + 1.8, y + 0.12, 4.4, 0.35, name, size=14, font=HEAD)
    text(s, L + 6.3, y + 0.12, W - 6.3, 0.35, intent, size=14)
    rule(s, L, y + rows_h, R, y + rows_h)
    t0 += mins

# =======================================================================================
# 3 — OUR IMPACT
# =======================================================================================
s = content("Our Impact: show the breadth, scale and value of what Confidence delivered", """
Intention: give the team a credible picture of the scale, diversity and importance of Confidence's contributions, using a few compelling metrics, short client stories and visual highlights. It should feel like one practice story, not a status update or a stack of team reports.

Thought starter: organize by the themes that came through in this year's Technology Reports rather than by client or reporting group.

Example stories (figures from the 2026 Technology Reports; validate before use):
• United Rentals: from experimenting with AI to the foundations of an AI-enabled delivery lifecycle. 90%+ automation coverage for Cart regression, 320+ automated tests for AU/NZ readiness, visual regression across 257 scenarios and 8 devices, Churchill AI-DLC expanding to more teams.
• Sherwin-Williams: Confidence standards became the operating model for the T4 transformation. Shared Agile, Jira and requirements standards across six T4 product teams; three releases in six weeks; AI-assisted requirements generating executable Playwright tests. (Reports cite five pods, six teams and seven workstreams at different points; reconcile before choosing a headline.)
• Mazda: clarity and confidence across an interconnected program. 30 Vehicle Landing Pages migrated, program-level AQ governance, the Open Intelligence Hub and Sidekick Agents. (Confirm the Oct 20 and 27 milestones before presenting them as complete.)

Show breadth beyond the largest accounts too. A visual summary or rotating collage can cover many highlights quickly.
""")
text(s, L, TOP, 5.3, 0.35, "Intention", size=16, font=HEAD)
text(s, L, TOP + 0.45, 5.3, 1.1, "A credible picture of our collective impact: a few strong metrics, short client stories and "
     "visual highlights that read as one practice story, not a status update.", size=14, line=1.1)
text(s, L, 3.75, 5.3, 0.35, "Thought starter: organize by theme, not by client", size=16, font=HEAD)
imp = ["Delivering real outcomes", "Reinventing delivery", "Building confidence at scale",
       "Establishing repeatable practices", "Growing as strategic partners"]
for i, h in enumerate(imp):
    y = 4.25 + i * 0.52
    text(s, L, y + 0.08, 0.55, 0.35, f"{i + 1:02d}", size=14, color=DGOLD, font=HEAD)
    text(s, L + 0.55, y + 0.08, 4.7, 0.35, h, size=14)
    rule(s, L, y + 0.5, L + 5.3, y + 0.5)
x2 = 6.05
text(s, x2, TOP, R - x2, 0.35, "Example stories", size=16, font=HEAD)
stories = [
    ("90%+", "United Rentals", "From AI experimentation to an AI-enabled delivery lifecycle"),
    ("6", "Sherwin-Williams", "Confidence standards became the operating model across T4 product teams"),
    ("30", "Mazda", "Vehicle Landing Pages migrated, with program-level AQ governance"),
]
for i, (stat, client, line_) in enumerate(stories):
    y = TOP + 0.5 + i * 1.0
    box(s, x2, y, R - x2, 0.88, line=BLACK)
    text(s, x2 + 0.2, y, 1.6, 0.88, stat, size=32, font=HEAD, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x2 + 1.85, y + 0.12, R - x2 - 2.05, 0.3, client, size=14, font=HEAD)
    text(s, x2 + 1.85, y + 0.44, R - x2 - 2.05, 0.4, line_, size=12, color=MGRAY, line=1.0)
text(s, x2, 5.55, R - x2, 0.35, "Breadth beyond the largest accounts", size=14, font=HEAD)
text(s, x2, 5.9, R - x2, 0.8, "Pokémon Automated Retail  ·  Tennessee Tourism  ·  Coca-Cola  ·  US Navy  ·  EchoPark  ·  "
     "Mayo Clinic  ·  Thrivent  ·  Global Payments  ·  Boehringer Ingelheim  ·  GDMS, Conagra, Spire",
     size=13, line=1.15)
text(s, x2, 6.7, R - x2, 0.3, "Examples from the 2026 Technology Reports; figures to be validated.", size=11, color=MGRAY)

# =======================================================================================
# 4 — OUR PEOPLE
# =======================================================================================
s = content("Our People: genuine recognition, tied to real contributions", """
Intention: keep the shout-outs that are traditionally part of the year-end All Hands, but connect them more directly to meaningful contributions. Reinforce that the year's impact came from people collaborating, adapting and investing in one another.

Thought starters:
• Recognize more than output: innovation and problem-solving, leadership through difficult delivery, collaboration, mentorship and knowledge sharing, improving our craft, and helping clients and colleagues succeed.
• Ask each group or pod lead for nominations with a short note: what they did, why it mattered, who benefited.
• Avoid a competitive awards format. A series of short spotlights will feel more authentic.
• Look past the reports so less visible contributors aren't missed.

The examples on the right come from the Technology Reports. They are starting points, not nominations.
""")
text(s, L, TOP, 4.6, 0.35, "Intention", size=16, font=HEAD)
text(s, L, TOP + 0.45, 4.6, 1.1, "Keep the shout-outs, but make them meaningful. The year's impact came from people "
     "collaborating, adapting and investing in one another.", size=14, line=1.1)
text(s, L, 3.75, 4.6, 0.35, "Thought starters", size=16, font=HEAD)
text(s, L, 4.2, 4.6, 2.8, [
    {"t": "Recognize growth, collaboration and resilience, not just output", "check": True},
    {"t": "Ask leads: what did they do, why did it matter, who benefited?", "check": True},
    {"t": "Short spotlights, not a competitive awards program", "check": True},
    {"t": "Look beyond the reports for less visible contributors", "check": True},
], size=14, after=8, line=1.05)
x2 = 5.3
text(s, x2, TOP, R - x2, 0.35, "Example spotlights from the Technology Reports", size=16, font=HEAD)
table(s, x2, TOP + 0.45, R - x2, [2.5, R - x2 - 2.5], [
    ("Theme", "Example"),
    ("AI-enabled delivery", "Kunal Pacheriwala: expanding the Churchill AI development platform"),
    ("Quality innovation", "Jacob Barrett: visual regression architecture; United Rentals automation team"),
    ("Program coordination", "Pan Leung: Navy workstream visibility and prioritization"),
    ("Complex delivery leadership", "Melissa Jones: Mazda cross-squad coordination and risk mitigation"),
    ("Practice standardization", "Sherwin-Williams T4 analyst team, incl. Abbey Logan on dependency management"),
    ("Knowledge sharing", "Spencer Brown: AI agent workshop and follow-up coaching"),
    ("Quality leadership", "Vanessa Lacy: Mazda program-level AQ; Kristy Riley: process modernization"),
    ("Community engagement", "Jessica Ward, Sharee Thompson and volunteers: KC Tech Council LEGiT program"),
], row_h=0.47, size=12)
text(s, x2, 6.75, R - x2, 0.3, "Starting points identified through reporting, not nominations.", size=11, color=MGRAY)

# =======================================================================================
# 5 — OUR FUTURE
# =======================================================================================
s = content("Our Future: a few provocations, then a live conversation about what we could become", """
Intention: encourage people to consider possibilities beyond familiar improvements and current priorities. This isn't a recap of the Future 100 or a 2027 strategy; it's an invitation to imagine.

The Confidence Future 3 (5 min): three original predictions inspired by VML Intelligence's Future 100: 2026, each connected to something we actually did in 2026.
1. AI as a delivery team member. Inspiration: Synthetic Generation, Democratizing creAItivity. 2026 connection: Churchill AI-DLC, AI-assisted requirements, Agent Ops.
2. Trust as our deliverable. Inspiration: Truth Literacy, Omnisurveillance, Coded Empathy. 2026 connection: continuous quality engineering and AQ governance.
3. Beyond the website. Inspiration: The Reality Interface, Acommerce, Digital Intent. 2026 connection: conversational AI and agent-enabled experiences.
A possible editorial format: a bold prediction, a short trend explanation, one example, a connection to our 2026 work, and one provocative question.

Popcorn conversation (12 min): an informal discussion where four or five people from different disciplines and pods react to a scenario and build on each other's ideas. "Yes, and," not evaluation; no consensus or business case required; the audience can join through chat. Kylie also floated involving two or three pods or small groups. Breakouts would allow broader participation but need more time.

Alternative theme: what if we reinvented Confidence from scratch? What would we keep, drop, reinvent or introduce?
""", dark=True)
fut = [
    ("What if AI becomes a member of every delivery team?",
     "What would we give an AI teammate, and what would we never delegate?"),
    ("What if trust becomes our most valuable deliverable?",
     "What if clients hired us primarily to establish trust in their technology?"),
    ("What if the website is no longer the destination?",
     "How do we define and validate experiences with no predictable journey?"),
]
text(s, L, TOP, W, 0.35, "The Confidence Future 3  ·  inspired by VML Intelligence, The Future 100: 2026",
     size=14, color=GRAY, font=HEAD)
cw = (W - 2 * 0.4) / 3
for i, (pred, q) in enumerate(fut):
    x = L + i * (cw + 0.4)
    text(s, x, TOP + 0.5, 0.8, 0.5, f"{i + 1:02d}", size=28, color=GOLD, font=HEAD)
    text(s, x, TOP + 1.1, cw - 0.1, 0.95, pred, size=19, color=WHITE, font=HEAD, line=1.05)
    text(s, x, TOP + 2.05, cw - 0.1, 0.7, [{"t": f"“{q}”", "italic": True}], size=13, color=LGOLD, line=1.1)
rule(s, L, 4.95, R, 4.95, color=MGRAY)
text(s, L, 5.15, 6.0, 0.35, "Then: a popcorn conversation", size=16, color=WHITE, font=HEAD)
text(s, L, 5.6, 6.0, 1.4, [
    {"t": "4–5 voices from different disciplines and pods build on one scenario", "check": True},
    {"t": "“Yes, and”: exploratory, not evaluative; the audience joins via chat", "check": True},
    {"t": "Kylie also floated two or three pods or small groups", "check": True},
], size=13, color=LGRAY, after=6, line=1.05)
x2 = 6.85
text(s, x2, 5.15, R - x2, 0.35, "Prompts to spark it", size=16, color=WHITE, font=HEAD)
text(s, x2, 5.6, R - x2, 1.4, [
    {"t": "If we were building Confidence from scratch today, what would we do differently?", "check": True},
    {"t": "Where could Technology Analysis, Scrum and Quality Engineering converge?", "check": True},
    {"t": "What would we pursue without today's roles, tools or processes?", "check": True},
], size=13, color=LGRAY, after=6, line=1.05)

prs.core_properties.title = "Confidence 2026: Our Impact, Our Future"
prs.core_properties.subject = "Year-End All Hands content outline"
prs.core_properties.author = "Zack Jemas"
prs.save(str(OUT))
print(f"Saved {OUT}  ({len(prs.slides._sldIdLst)} slides)")
