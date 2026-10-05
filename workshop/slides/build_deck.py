"""Builds rag-workshop.pptx, the slides for the 90-minute workshop.

Run from this folder:  uv run --with python-pptx --with "qrcode[pil]" python build_deck.py
Needs rsvg-convert (brew install librsvg) to turn ../img/*.svg into pictures.
For the PDF, open the .pptx in PowerPoint: File > Export > PDF.
Layout and images follow the KISZ deck from the Brandenburger Digitaltag 2026.
"""
import subprocess
import tempfile
from pathlib import Path

import qrcode
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

HERE = Path(__file__).parent
ASSETS = HERE / "assets"
TMP = Path(tempfile.mkdtemp())
REPO = "github.com/aihpi/workshop-rag-basics"

RED, ORANGE, YELLOW, BLUE = "B1063A", "DD6108", "F6A800", "007A9E"
GREY, MUTED, INK, DARK, SAND = "5A6065", "8B9094", "17161C", "30343A", "EDEBE7"
WHITE = "FFFFFF"

prs = Presentation()
prs.slide_width, prs.slide_height = Emu(18288000), Emu(10477500)  # 20 × 11.46 in, as the original
W, H = 20.0, 10477500 / 914400
BLANK = prs.slide_layouts[6]


# --------------------------------------------------------------------------- helpers
def rgb(hex_):
    return RGBColor.from_string(hex_)


def text(slide, x, y, w, h, runs, size=18, color=INK, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, font="Arial", spacing=0, line=1.15, rotation=0):
    """runs: a string, or a list of paragraphs; a paragraph is a string or a list of (text, {overrides})."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    box.rotation = rotation
    frame = box.text_frame
    frame.word_wrap, frame.vertical_anchor = True, anchor
    frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
    for i, para in enumerate([runs] if isinstance(runs, str) else runs):
        p = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        p.alignment, p.line_spacing = align, line
        for run_text, opts in ([(para, {})] if isinstance(para, str) else para):
            r = p.add_run()
            r.text = run_text
            f = r.font
            f.name, f.size = opts.get("font", font), Pt(opts.get("size", size))
            f.bold, f.color.rgb = opts.get("bold", bold), rgb(opts.get("color", color))
            if opts.get("link"):
                r.hyperlink.address = opts["link"]
            if opts.get("spacing", spacing):
                r._r.get_or_add_rPr().set("spc", str(int(opts.get("spacing", spacing) * 100)))
        if isinstance(para, list) and para and para[0][1].get("space_after"):
            p.space_after = Pt(para[0][1]["space_after"])
    return box


def rect(slide, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(fill)
    if line:
        s.line.color.rgb, s.line.width = rgb(line), Pt(1.5)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    return s


def poly(slide, points, fill):
    builder = slide.shapes.build_freeform(Inches(points[0][0]), Inches(points[0][1]), scale=Inches(1))
    builder.add_line_segments([(x - points[0][0], y - points[0][1]) for x, y in points[1:]])
    s = builder.convert_to_shape(Inches(points[0][0]), Inches(points[0][1]))
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(fill)
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def line(slide, points, color, width=2.0, arrow=False):
    for (x1, y1), (x2, y2) in zip(points, points[1:]):
        c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        c.line.color.rgb, c.line.width = rgb(color), Pt(width)
    if arrow:
        c.line._get_or_add_ln().append(c.line._ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))


def image(slide, name, x, y, w=None, h=None):
    path = name if isinstance(name, Path) else ASSETS / name
    return slide.shapes.add_picture(str(path), Inches(x), Inches(y),
                                    Inches(w) if w else None, Inches(h) if h else None)


def diagram(name, x, y, w, slide):
    png = TMP / f"{name}.png"
    if not png.exists():
        subprocess.run(["rsvg-convert", "-z", "3", "-b", "white", str(HERE.parent / "img" / f"{name}.svg"), "-o", str(png)], check=True)
    return image(slide, png, x, y, w=w)


def qr(url, name):
    path = TMP / f"{name}.png"
    qrcode.make(url, border=1).save(path)
    return path


def notes(slide, body):
    slide.notes_slide.notes_text_frame.text = body


def page_number(slide):
    text(slide, 19.3, 10.9, 0.5, 0.35, str(len(prs.slides)), size=16, color=MUTED, align=PP_ALIGN.RIGHT)


def logos(slide):
    image(slide, "logo-bmftr.png", 13.94, 0.35, w=2.62)
    image(slide, "logo-kisz.png", 16.77, 0.43, w=2.8)


def content(title, band, label, logo=True, background=None):
    """White slide with the coloured band on the left, a vertical section label and the title."""
    slide = prs.slides.add_slide(BLANK)
    if background:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = rgb(background)
    poly(slide, [(0, 0), (0.94, 0), (0.94, 8.85), (0, H)], band)
    if label:
        width = 0.36 * len(label) + 0.6
        text(slide, 0.47 - width / 2, 6.0 - 0.18, width, 0.36, label, size=15, bold=True, color=WHITE if band != YELLOW else INK,
             spacing=4, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, rotation=270)
    text(slide, 1.16, 0.43, 12.5, 1.4, title.upper(), size=39, bold=True, color=WHITE if background else GREY,
         anchor=MSO_ANCHOR.MIDDLE, line=1.0)
    if logo:
        logos(slide)
    page_number(slide)
    return slide


def divider(number, title, subtitle, fill, ink=WHITE):
    slide = prs.slides.add_slide(BLANK)
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(fill)
    text(slide, 1.16, 2.85, 12, 0.5, number, size=22, bold=True, color=ink, spacing=3)
    text(slide, 1.16, 3.45, 17, 2.3, title.upper(), size=60, bold=True, color=ink, line=0.95, anchor=MSO_ANCHOR.BOTTOM)
    rect(slide, 1.16, 6.0, 1.75, 0.08, ink)
    text(slide, 1.16, 6.45, 16, 0.6, subtitle, size=24, color=ink)
    poly(slide, [(0, 8.85), (4.65, H), (0, H)], WHITE)
    image(slide, "logo-bmftr.png", 0.05, 9.75, w=2.62)
    page_number(slide)
    return slide


def numbered(slide, x, y, w, items, colour=YELLOW, size=18, gap=0.72):
    for i, item in enumerate(items, start=1):
        rect(slide, x, y + (i - 1) * gap, 0.42, 0.42, colour)
        text(slide, x, y + (i - 1) * gap, 0.42, 0.42, str(i), size=16, bold=True, color=WHITE if colour != YELLOW else INK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(slide, x + 0.65, y + (i - 1) * gap + 0.03, w - 0.65, gap, item, size=size)


def kicker(slide, x, y, w, words, color=MUTED):
    text(slide, x, y, w, 0.4, words.upper(), size=15, bold=True, color=color, spacing=3)


def chat(slide, x, y, w, h, exchange):
    """A chat window drawn from shapes: [(who, text, extras)], who is 'you', 'bot' or 'step'."""
    rect(slide, x, y, w, h, "F4F3F1", shape=MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.04
    text(slide, x + 0.35, y + 0.25, w - 0.7, 0.35, "localhost:8000", size=13, color=MUTED, font="Courier New")
    cy = y + 0.85
    for who, body, height in exchange:
        if who == "you":
            bw = min(w - 1.6, 0.13 * len(body) + 0.6)
            rect(slide, x + w - 0.35 - bw, cy, bw, height, "E2E0DC", shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            text(slide, x + w - 0.2 - bw, cy + 0.12, bw - 0.3, height - 0.2, body, size=15)
        elif who == "step":
            text(slide, x + 0.35, cy, w - 0.7, height, [[("✓ ", {"color": ORANGE, "bold": True}), (body, {"color": GREY})]], size=15)
        else:
            rect(slide, x + 0.35, cy, w - 0.7, height, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.08
            text(slide, x + 0.6, cy + 0.15, w - 1.2, height - 0.25, body, size=15)
        cy += height + 0.25


def notebook(number, title, minutes, learn, turn, exchange):
    slide = content(title, YELLOW, "HANDS-ON")
    kicker(slide, 1.16, 1.85, 9, f"Notebook {number} · {minutes} min", ORANGE)
    kicker(slide, 1.16, 2.65, 9, "In this notebook")
    text(slide, 1.16, 3.15, 9.4, 2.4, [[(f"•  {item}", {"space_after": 8})] for item in learn], size=19)
    kicker(slide, 1.16, 5.55, 9, "Your turn · in the chatbot")
    numbered(slide, 1.16, 6.05, 9.4, turn, gap=0.75)
    text(slide, 1.16, 6.05 + 0.75 * len(turn) + 0.15, 9.4, 0.4, "Solutions are folded away under each exercise.",
         size=15, color=MUTED)
    chat(slide, 11.2, 2.3, 8.3, 1.1 + sum(h + 0.25 for _, _, h in exchange), exchange)
    return slide


# ---------------------------------------------------------------------------- 1 title
s = prs.slides.add_slide(BLANK)
text(s, 1.67, 2.24, 15, 0.4, "[EVENT]", size=19.5, bold=True, color="DC640D", spacing=4)
text(s, 1.67, 2.82, 15, 1.1, "BUILD A RAG CHATBOT", size=69, bold=True, color=GREY)
text(s, 1.67, 4.11, 15, 0.7, "AI-assisted research on your own documents", size=34.5, color=RED)
rect(s, 1.67, 5.2, 1.25, 0.08, RED)
text(s, 1.67, 5.76, 14, 1.1, [[("[Presenter]", {"size": 22.5})],
                              [("AI Service Centre Berlin-Brandenburg  |  [Date]", {"size": 22.5, "color": GREY})]])
image(s, "logo-kisz.png", 16.45, 0.74, w=2.8)
image(s, "logo-bmftr.png", 16.41, 1.98, w=2.8)
image(s, "hpi-tagline.png", 0.76, 10.23, w=2.92)
image(s, "hpi-building.png", 9.72, 6.72, w=11.72)
notes(s, "Welcome. 90 minutes. Goal: understand how AI-assisted research on your own documents works, "
         "and build it yourself, step by step, in five notebooks.")

# ---------------------------------------------------------------------------- 2 participant list
s = content("Participant list", MUTED, None)
text(s, 1.16, 2.6, 10, 1.2, "The list is going round now. Please sign in while everyone is still seated.", size=24)
for i, head in enumerate(["Name", "Institution", "Email"]):
    rect(s, 1.16 + i * 3.13, 4.3, 3.1, 0.75, ORANGE)
    text(s, 1.36 + i * 3.13, 4.3, 2.8, 0.75, head, size=19, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    for row in range(3):
        rect(s, 1.16 + i * 3.13, 5.08 + row * 0.6, 3.1, 0.57, SAND)
rect(s, 12.3, 2.75, 6.25, 0.08, MUTED)
text(s, 12.3, 3.1, 6.25, 1.2, "The printout stays with the workshop lead. Please hand it back at the end.", size=19, color=GREY)
notes(s, "Pass the list round now, not at the end, or it gets lost on the way. Collect it again at the end.")

# ---------------------------------------------------------------------------- agenda
s = content("Agenda · 90 minutes", MUTED, None)
AGENDA = [(0, 10, RED, "Slides", "Welcome, the AI Service Centre, why RAG"),
          (10, 5, BLUE, "Slides", "Architecture and setup"),
          (15, 10, YELLOW, "Notebook 1", "Talking to an LLM"),
          (25, 13, YELLOW, "Slides + notebook 2", "From document to chunks · reading documents"),
          (38, 7, BLUE, "Slides", "What we found so far · embeddings and vector search"),
          (45, 13, YELLOW, "Notebook 3", "Search"),
          (58, 16, YELLOW, "Slides + notebook 4", "Prompt and citations · retrieval-augmented generation"),
          (74, 10, YELLOW, "Notebook 5", "Putting it together"),
          (84, 6, RED, "Slides", "Where to go next, feedback")]
for i, (start, minutes, colour, kind, what) in enumerate(AGENDA):
    y = 2.05 + i * 0.74
    text(s, 1.16, y, 1.2, 0.6, f"{start // 60}:{start % 60:02d}", size=20, bold=True, color=GREY, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, 2.5, y + 0.12, 0.36, 0.36, colour)
    text(s, 3.15, y, 4.2, 0.6, kind, size=18, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 7.4, y, 9.8, 0.6, what, size=18, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 17.3, y, 1.55, 0.6, f"{minutes} min", size=17, color=MUTED, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
per_minute = 17.7 / 90
for start, minutes, colour, _, _ in AGENDA:
    rect(s, 1.16 + start * per_minute, 9.0, minutes * per_minute - 0.04, 0.42, colour)
hands_on = 10 + 10 + 13 + 13 + 10  # the notebook times; the mixed blocks open with 3 min of slides
text(s, 1.16, 9.6, 17.7, 0.5, [[("■ ", {"color": YELLOW}), (f"notebooks, hands-on: {hands_on} min     ", {}),
                                ("■ ", {"color": BLUE}), ("theory     ", {}), ("■ ", {"color": RED}), ("welcome and wrap-up", {})]],
     size=16, color=GREY)
notes(s, "The plan for today: theory in small doses, most of the time in the notebooks. "
         "The two mixed blocks start with 3 minutes of slides, then the notebook.")

# ---------------------------------------------------------------------------- 3–7 KISZ
s = divider("01", "AI Service Centre\nBerlin-Brandenburg", "Who we are and what you can use.", RED)
notes(s, "Transition: who we are, and how you can keep working with us. About 5 minutes for slides 3 to 7.")

s = content("Four AI service centres", BLUE, "AI SERVICE CENTRES")
image(s, "qr-kisz.png", 1.25, 2.75, w=2.85)
text(s, 1.25, 5.65, 2.85, 0.35, "hpi.de/ki-servicezentrum", size=13, color=GREY, align=PP_ALIGN.CENTER)
image(s, "map-germany.png", 5.46, 1.85, w=5.21)
image(s, "map-states.png", 5.46, 1.98, w=5.19)
centres = [  # colour, dots on the map, path to the label, logo, place, focus
    ("1BABE3", [(7.75, 4.08), (7.8, 4.88), (7.5, 5.2)], [(7.99, 4.88), (8.47, 4.88), (10.82, 2.8), (13.39, 2.8)],
     "logo-kisski.png", "Göttingen | Kassel | Hannover", "Sensitive and critical infrastructure, medicine & energy"),
    (ORANGE, [(9.55, 4.29)], [(9.55, 4.29), (13.39, 4.29)],
     "logo-kisz-small.png", "Potsdam | Berlin", "Education and consulting, AI in business & society"),
    ("E9358D", [(6.37, 4.99), (6.42, 5.25), (5.92, 5.59), (5.74, 5.7), (6.52, 5.97)], [(6.45, 5.05), (7.17, 5.52), (13.39, 5.52)],
     "logo-westai.png", "Bonn | St. Augustin | Aachen | Jülich | Dortmund", "AI hardware, consulting and training; multimodal and transferable models"),
    ("8A6AF4", [(6.97, 6.24)], [(7.17, 6.24), (9.6, 6.24), (10.8, 6.8), (13.39, 6.8)],
     "logo-hessian.png", "Darmstadt", "Explainability, generalisability and contextual adaptation"),
]
for i, (colour, dots, path, logo, place, focus) in enumerate(centres):
    for dx, dy in dots:
        rect(s, dx - 0.06, dy - 0.06, 0.12, 0.12, colour, shape=MSO_SHAPE.OVAL)
    line(s, path, colour, 2.25)
    rect(s, path[-1][0] - 0.1, path[-1][1] - 0.1, 0.2, 0.2, colour, shape=MSO_SHAPE.OVAL)
    top = [2.35, 3.8, 5.05, 6.4][i]
    image(s, logo, 11.85, top, w=1.35)
    text(s, 13.7, top - 0.05, 6.2, 0.4, place, size=19, bold=True, color=colour)
    text(s, 13.7, top + 0.4, 6.0, 0.8, focus, size=16)
image(s, "logo-centres.png", 6.01, 9.47, w=4.48)
text(s, 15.2, 8.9, 4.6, 1.8, [[("Goal: ", {"bold": True}), ("lower the barriers to putting AI to work in society and business", {"bold": True})]], size=18)
notes(s, "Four centres across Germany. Our site in Potsdam and Berlin covers education, consulting "
         "and the use of AI in business and society.")

s = content("Education", YELLOW, None, logo=False)
image(s, "qr-newsletter.png", 1.16, 2.4, w=2.85)
text(s, 1.16, 5.3, 2.85, 0.35, "Newsletter", size=15, bold=True, color=GREY, align=PP_ALIGN.CENTER)
rows = [  # colour, label, qr, link, bullets, tiles
    (YELLOW, "Talks", "qr-talks.png", "tele-task.de/series/1463", ["Guest talks on research and innovation"],
     ["tile-talk-1.png", "tile-talk-2.png", "tile-talk-3.jpeg"]),
    (ORANGE, "Work-\nshops", "qr-workshops.png", "aimaker.community",
     ["Hands-on topics", "For example: speech2summary, Docker for ML, semantic search"],
     ["tile-workshop-1.jpeg", "tile-workshop-2.jpeg", "tile-workshop-3.jpeg"]),
    (RED, "MOOCs", "qr-moocs.png", "open.hpi.de/channels/ai-service-center",
     ["ChatGPT: what does generative AI mean for society?", "Profitable AI", "Understanding and avoiding AI bias"],
     ["tile-mooc-1.jpeg", "tile-mooc-2.png", "tile-mooc-3.jpeg"]),
]
for i, (colour, label, code, link, bullets, tiles) in enumerate(rows):
    top = 0.55 + i * 3.6
    rect(s, 5.32, top, 1.75, 1.7, colour)
    text(s, 5.32, top, 1.75, 1.7, label, size=21, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line=1.0)
    image(s, code, 9.0, top, w=1.5)
    text(s, 7.3, top + 1.55, 3.4, 0.35, link, size=13, bold=True, color="DC640D", align=PP_ALIGN.RIGHT)
    text(s, 5.32, top + 1.95, 5.6, 1.5, [[(f"•  {b}", {"space_after": 2})] for b in bullets], size=16)
    for j, tile in enumerate(tiles):
        image(s, tile, 11.0 + j * 3.0, top, w=2.82)
notes(s, "Everything is free and open. Newsletter via the QR code on the left.")

s = content("Infrastructure", DARK, None, logo=False)
image(s, "logo-kisz.png", 16.77, 0.43, w=2.8)
image(s, "qr-aisc.png", 1.6, 2.75, w=2.85)
text(s, 1.6, 5.65, 2.85, 0.35, "aisc.hpi.de · request access", size=13, color=GREY, align=PP_ALIGN.CENTER)
image(s, "compute-banner.jpeg", 5.09, 2.0, w=7.0)
image(s, "compute-racks.jpeg", 5.09, 5.95, w=7.0, h=5.2)
O = {"color": ORANGE}
terms = [[("•  Access ", {}), ("free of charge", O)], [("•  No production use", {})],
         [("      ◦  Data should be ", {}), ("anonymised or synthetic", O)], [("      ◦  No ", {}), ("hosting of products", O)],
         [("•  ", {}), ("Reporting & publication", O), (" by users", {})], [("•  ", {}), ("Existing rights", O), (" stay with users", {})],
         [("•  ", {}), ("New rights", O), (" stay with users", {})], [("      ◦  Usage rights granted for research and teaching", {})]]
text(s, 12.6, 1.95, 7.2, 4.2, terms, size=17, line=1.25)
specs = [("Training", ["64 NVIDIA H100 GPUs"]), ("Inference", ["40 NVIDIA A30 GPUs"]),
         ("ARM server", ["Ampere Altra Max M128-30 CPU", "2× NVIDIA L40 GPUs"]), ("GPU server", ["AMD Epyc CPU", "8× NVIDIA L40S GPUs"]),
         ("Edge", ["ARMv8 CPU", "NVIDIA Jetson AGX modules"]), ("Neuromorphic", ["288 SpiNNaker2 chips"]),
         ("Storage", ["1.5 PB NVRAM"]), ("Network", ["400 Gb/s InfiniBand", "200 Gb/s Ethernet"])]
for col, chunk in enumerate([specs[:4], specs[4:]]):
    paras = []
    for head, items in chunk:
        paras.append([(head, {"bold": True, "color": ORANGE})])
        paras += [[(f"•  {item}", {})] for item in items]
        paras.append([(" ", {"size": 6})])
    text(s, 12.6 + col * 3.7, 6.0, 3.6, 4.9, paras, size=16, line=1.1)
notes(s, "Free to use, no production use. Request access at aisc.hpi.de.")

s = content("Consulting", ORANGE, None, logo=False)
image(s, "qr-office-hours.png", 1.16, 2.55, w=2.95)
text(s, 1.16, 5.55, 2.95, 0.35, "Book office hours", size=13, bold=True, color=GREY, align=PP_ALIGN.CENTER)
blocks = [(RED, "AI office hours", ["AI infrastructure", "AI models & frameworks", "AI use cases"]),
          (ORANGE, "AI pilot projects", ["Co-developing a prototype", "Applications every three months",
                                         "Selection criteria such as AI maturity and public benefit", "Results are published"]),
          (YELLOW, "Collaborations", ["Jointly organised network meetings"]),
          (INK, "Past AI pilot projects", ["Generating maths problems", "Generating plain language", "Generating upcycling ideas",
                                           "Reducing food waste", "Dating documents by handwriting"])]
paras = []
for colour, head, items in blocks:
    paras.append([(head, {"bold": True, "color": colour, "size": 21})])
    paras += [[(f"•  {item}", {})] for item in items]
    paras.append([(" ", {"size": 8})])
text(s, 4.9, 2.2, 7.8, 8.8, paras, size=18, line=1.1)
image(s, "consulting.jpeg", 12.95, 0.3, w=6.6)
image(s, "qr-apply.png", 16.15, 7.0, w=2.95)
text(s, 16.15, 10.0, 2.95, 0.35, "Apply now!", size=13, bold=True, color="DC640D", align=PP_ALIGN.CENTER)
notes(s, "Concrete call to action: book office hours (QR on the left), apply for a pilot project, "
         "which opens every three months (QR on the right).")

# ---------------------------------------------------------------------------- 8–10 motivation
s = divider("02", "Motivation", "What language models don't know, and what that means for research.", RED)
notes(s, "Core question: why is a language model on its own not enough for research in specialist documents? About 5 minutes for 8 to 10.")

s = content("Why a language model\nalone is not enough", RED, "MOTIVATION")
cols = [(ORANGE, "No access to\nyour own data",
         "Language models are trained mostly on public data. Internal documents about your processes and paywalled "
         "specialist publications are unknown to them."),
        (RED, "Similar instead\nof right",
         "When the right information is missing, the model falls back on related material: the rules for no-parking "
         "zones answer a question about no-stopping zones, popular science stands in for the specialist literature."),
        (YELLOW, "RAG decides what\nthe answer is built on",
         "Retrieval-augmented generation sets which documents an answer may use, and shows which source every "
         "statement comes from.")]
for i, (colour, head, body) in enumerate(cols):
    x = 1.16 + i * 6.12
    rect(s, x, 2.5, 5.46, 0.08, colour)
    text(s, x, 2.85, 5.5, 0.4, f"0{i + 1}", size=19, bold=True, color=MUTED, spacing=3)
    text(s, x, 3.35, 5.5, 1.2, head, size=28, bold=True, line=1.05)
    text(s, x, 4.75, 5.3, 4.5, body, size=21, line=1.3)
notes(s, "The no-parking versus no-stopping example makes the point fastest: the model answers plausibly, "
         "but from the wrong legal text. You'll see the same in notebook 1, with a scientific paper.")

s = content("RAG in practice", ORANGE, "IN PRACTICE")
for i, (shot, name, what, code) in enumerate([
        ("screenshot-sentra.png", "SENTRA", "Research Services of the German Bundestag: semantic search across "
         "expert reports, an AI summary, every statement traced to its report.", "github.com/aihpi/pilotproject-sentra"),
        ("screenshot-grundschutzki.png", "GrundschutzKI", "IT baseline protection: questions in plain language, "
         "a short answer with a jump link, the BSI original open next to it.", "github.com/aihpi/pilotprojekt-GrundschutzKI")]):
    x = 1.16 + i * 9.35
    pic = image(s, shot, x, 2.2, w=8.7)
    if pic.height > Inches(4.6):  # keep both screenshots the same height
        pic.height, pic.width = Inches(4.6), int(pic.width * Inches(4.6) / pic.height)
    text(s, x, 7.05, 8.7, 0.5, name, size=24, bold=True)
    text(s, x, 7.6, 8.5, 1.3, what, size=18, line=1.25)
    text(s, x, 8.95, 8.7, 0.4, code, size=15, color="DC640D", font="Courier New")
text(s, 1.16, 9.8, 17.5, 0.5, [[("Semantic search  ·  an AI summary  ·  sources for transparency.  ", {"bold": True}),
                                 ("Today you build the same thing, small.", {"color": GREY})]], size=19)
notes(s, "No live demo today, just these two to show RAG in real use. SENTRA: up to 2000 requests a year at the "
         "Bundestag's research services. GrundschutzKI: do-it-yourself instead of external consulting, also with NIS2 in mind.")

# ---------------------------------------------------------------------------- 11–13 build it yourself
s = divider("03", "Build it yourself", "Five notebooks, one chatbot. A little theory before each.", YELLOW, ink=INK)
notes(s, "From here: theory in small doses, then hands-on. Write the gateway key on the whiteboard.")

s = content("Architecture", BLUE, "BASICS")
rect(s, 1.32, 2.2, 4.05, 6.1, "EAF4F8", line="9CCBDD")
rect(s, 5.52, 2.2, 9.63, 6.1, "FDF1E9", line="F2B994")
for x, n, colour, head, sub in [(1.55, "1", BLUE, "INGESTION", "once · notebooks 2 + 3"),
                                (5.75, "2", ORANGE, "INFERENCE", "every question · notebooks 3 + 4")]:
    rect(s, x, 2.42, 0.42, 0.42, colour)
    text(s, x, 2.42, 0.42, 0.42, n, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + 0.6, 2.38, 6, 0.45, head, size=19, bold=True, color=GREY, spacing=3)
    text(s, x + 0.6, 2.82, 7, 0.4, sub, size=15, color=MUTED)
boxes = {  # x, y, w, fill, ink, title, subtitle
    "docs": (2.0, 3.36, 3.15, "E7E5E1", INK, "Documents", "PDF, Word, HTML"),
    "embed": (2.0, 5.05, 5.4, "E7E5E1", INK, "Embedding model", "Text in, vector out"),
    "db": (2.0, 6.72, 5.4, BLUE, WHITE, "Vector database", "Index for similarity search"),
    "ask": (8.23, 3.36, 3.17, ORANGE, WHITE, "Question", "asked by the user"),
    "hits": (8.23, 6.72, 3.17, "E7E5E1", INK, "Relevant passages", "go into the prompt"),
    "llm": (11.98, 5.05, 3.13, "E7E5E1", INK, "LLM", "language model"),
    "answer": (15.69, 5.05, 3.17, ORANGE, WHITE, "Answer", "with sources"),
}
for x, y, w, fill, ink, head, sub in boxes.values():
    rect(s, x, y, w, 1.29, fill)
    text(s, x + 0.27, y + 0.22, w - 0.4, 0.45, head, size=20, bold=True, color=ink)
    text(s, x + 0.27, y + 0.7, w - 0.4, 0.4, sub, size=16, color=ink)
line(s, [(3.57, 4.65), (3.57, 5.0)], BLUE, 3, arrow=True)
line(s, [(3.57, 6.34), (3.57, 6.67)], BLUE, 3, arrow=True)
line(s, [(8.35, 4.77), (7.5, 5.62)], ORANGE, 3, arrow=True)
line(s, [(6.04, 6.34), (6.04, 6.67)], ORANGE, 3, arrow=True)
line(s, [(7.4, 7.36), (8.18, 7.36)], ORANGE, 3, arrow=True)
line(s, [(11.25, 4.77), (11.92, 5.47)], ORANGE, 3, arrow=True)
line(s, [(11.25, 6.67), (11.92, 5.97)], ORANGE, 3, arrow=True)
line(s, [(15.14, 5.7), (15.64, 5.7)], ORANGE, 3, arrow=True)
text(s, 1.16, 8.85, 8.4, 2.0, [[("Once, ahead of time, ", {"bold": True, "color": BLUE}),
                                ("the documents are read and cut into chunks. The embedding model turns every chunk into a "
                                 "vector, a list of numbers that captures its meaning. The vectors go into the vector database.", {})]], size=17, line=1.25)
text(s, 10.3, 8.85, 8.6, 2.0, [[("For every question, ", {"bold": True, "color": ORANGE}),
                                ("the same model turns the question into a vector and searches the same database. The closest "
                                 "chunks go to the language model with the question, and it answers with sources.", {})]], size=17, line=1.25)
notes(s, "3 minutes. Stress the two phases: ingestion happens once, inference with every question. "
         "The same embedding model in both. The badges show which notebook builds which part.")

s = content("Setup", YELLOW, "HANDS-ON")
steps = [("Get the code", "git clone https://github.com/aihpi/workshop-rag-basics.git\ncd workshop-rag-basics/workshop"),
         ("Install", "uv sync"), ("Add your key", "cp .env.example .env    # then paste the key"),
         ("Check", "uv run python check_setup.py"), ("Start", "uv run jupyter lab")]
for i, (head, code) in enumerate(steps):
    y = 2.15 + i * 1.32 + (0.32 if i else 0)
    rect(s, 1.16, y, 0.42, 0.42, YELLOW)
    text(s, 1.16, y, 0.42, 0.42, str(i + 1), size=16, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 1.85, y - 0.02, 4, 0.45, head, size=19, bold=True)
    text(s, 1.85, y + 0.45, 9.5, 0.8, code, size=15, font="Courier New", color=GREY, line=1.2)
image(s, qr("https://" + REPO, "qr-repo"), 9.3, 2.1, w=1.7)
text(s, 8.3, 3.85, 3.7, 0.35, REPO.split("/", 1)[1], size=12, color=GREY, align=PP_ALIGN.CENTER)
diagram("live-reload", 12.0, 2.2, 7.6, s)
text(s, 12.0, 4.3, 7.4, 1.8, "Cells that start with %%writefile app.py write the chatbot. Chainlit notices the change and "
     "reloads it in your browser at localhost:8000.", size=18, line=1.25)
for i, (head, value) in enumerate([("Chat model", "gemma-4-31b"), ("Embedding model", "octen-embedding-8b"),
                                   ("Runs on", "our own infrastructure in Potsdam")]):
    kicker(s, 12.0, 6.55 + i * 1.15, 7, head)
    text(s, 12.0, 6.95 + i * 1.15, 7.4, 0.45, value, size=19, font="Courier New" if i < 2 else "Arial")
notes(s, "2 minutes, if everyone ran check_setup.py at home. Whoever didn't: start uv sync and check_setup.py now; "
         "Docling downloads about 500 MB of models the first time. Notebook 1 works before that finishes.")

# ---------------------------------------------------------------------------- 14–23 notebooks with theory in between
s = notebook(1, "Talking to an LLM", 10,
             ["Connect to the model through the gateway", "Ask about a paper the model has never seen",
              "Build a chatbot that remembers and streams"],
             ["Personality: change the system prompt", "Starter buttons with @cl.set_starters",
              "Stopwatch: show how long each answer took"],
             [("you", "In Kage et al. (2018), what were the lifetime-encoded beads loaded with?", 1.0),
              ("bot", "The beads were loaded with quantum dots (QDs), with lifetimes of about 2.5 ns, 10 ns and 25 ns.", 1.3),
              ("step", "Sounds sure. It's wrong: the paper used organic dyes and quantum dots, with lifetimes from 1.7 to 22.6 ns.", 1.0)])
notes(s, "10 minutes. Point out the confident wrong answer, it comes back in notebook 4. Then the Your turn exercises.")

s = content("From document to chunks", BLUE, "BASICS")
kicker(s, 1.16, 2.0, 8, "Step 1 · Docling reads the page")
diagram("docling", 1.16, 2.5, 8.6, s)
text(s, 1.16, 6.3, 8.4, 1.4, "A PDF only stores where to draw each letter. Docling finds headings, paragraphs, tables and "
     "figures, and writes Markdown.", size=18, line=1.25)
kicker(s, 10.5, 2.0, 8, "Step 2 · cut into chunks")
diagram("chunking", 10.5, 2.5, 8.6, s)
text(s, 10.5, 6.3, 8.6, 4, [[("Our notebook: ", {"bold": True}), ("every 1000 characters, with 200 overlap. Simple, the same "
                             "for every document, but it may cut through a table.", {})],
                            [(" ", {"size": 8})],
                            [("GrundschutzKI: ", {"bold": True}), ("at every heading, because the BSI standards are clearly "
                             "structured. A chunk keeps its title. Docling ships several such chunkers.", {})]], size=18, line=1.25)
notes(s, "3 minutes. One method of many: fixed size with overlap is what the notebook does. "
         "Structured documents chunk better along headings.")

s = notebook(2, "Reading documents", 10,
             ["Convert the three papers with Docling, page by page", "See a table come through as a table",
              "Cut the pages into chunks, and look where the cuts land"],
             ["Where are the cuts? Send the first three chunks", "Only the tables: send every table Docling found",
              "If time: show the original PDF next to it"],
             [("you", "📎 Schmidt_2022_SciReports.pdf", 0.62),
              ("step", "Docling reads Schmidt_2022_SciReports.pdf", 0.45),
              ("bot", "Schmidt_2022_SciReports.pdf: 10 pages, 54 chunks. This is page 1:\n\nA multiparametric fluorescence "
                      "assay for screening aptamer-protein interactions based on microbeads …", 2.1)])
notes(s, "10 minutes. Docling takes about a minute for the three papers; talk through the Markdown while it runs.")

s = content("What we found so far", YELLOW, "CHECKPOINT", logo=False, background=GREY)
rect(s, 13.0, 0.85, 4.9, 1.1, YELLOW)
text(s, 13.25, 0.85, 4.5, 1.1, "Ask the room first,\nthen show the finding.", size=17, anchor=MSO_ANCHOR.MIDDLE)
FINDINGS = [("Notebook 1", "Did the model know the Kage paper?",
             "No, but it answered anyway: quantum dots, 2.5, 10 and 25 ns. The paper used organic dyes and quantum dots, "
             "1.7 to 22.6 ns. A model doesn't know your documents, and doesn't say so."),
            ("Notebook 2", "What did Docling do with the tables?",
             "It kept all 5 tables in the three papers as Markdown tables, with their rows and columns, ready to be "
             "searched like any other text."),
            ("Notebook 2 · your turn", "How big should a chunk be?", None)]
for i, (where, question, finding) in enumerate(FINDINGS):
    x = 1.16 + i * 6.12
    rect(s, x, 2.5, 5.46, 0.08, YELLOW)
    text(s, x, 2.85, 5.5, 0.4, where.upper(), size=15, bold=True, color="D0D3D6", spacing=3)
    text(s, x, 3.35, 5.4, 1.3, question, size=26, bold=True, color=WHITE, line=1.05)
    rect(s, x, 4.85, 5.46, 3.75, WHITE)
    if finding:
        text(s, x + 0.3, 5.1, 4.9, 3.3, finding, size=20, line=1.3)
        continue
    for r, (size, count, cut) in enumerate([("Size", "Chunks", "End mid-table"), ("300", "628", "17"),
                                            ("1000", "197", "5"), ("3000", "69", "0")]):
        bold = r == 0
        for c, value in enumerate((size, count, cut)):
            text(s, x + 0.3 + c * 1.55, 5.1 + r * 0.5, 1.6 if c < 2 else 2.0, 0.45, value, size=16 if bold else 18,
                 bold=bold, color=MUTED if bold else INK)
    text(s, x + 0.3, 7.25, 4.9, 1.3, "Small chunks find the exact passage but lose context; big ones keep tables whole "
         "but blur the match.", size=15, line=1.25, color=GREY)
rect(s, 1.16, 9.0, 17.7, 1.1, WHITE)
text(s, 1.5, 9.0, 3, 1.1, "NEXT", size=15, bold=True, color=MUTED, spacing=3, anchor=MSO_ANCHOR.MIDDLE)
text(s, 3.4, 9.0, 15.2, 1.1, "We have good chunks. How do we find the right ones for a question? Embeddings.",
     size=20, anchor=MSO_ANCHOR.MIDDLE)
notes(s, "4 minutes. Ask each question to the room before pointing at the answer. The numbers are real: "
         "the notebook 1 answer from a test run (it varies a little), the chunk counts from the three papers.")

s = content("Embeddings and vector search", BLUE, "BASICS")
diagram("embeddings", 1.16, 2.1, 7.4, s)
text(s, 1.16, 6.85, 7.2, 2.5, [[("An embedding model ", {"bold": True}), ("turns text into 4096 numbers. Similar meaning, "
                               "similar direction, even without shared words.", {})]], size=18, line=1.25)
diagram("qdrant", 8.95, 2.1, 10.6, s)
text(s, 8.95, 6.85, 10.4, 2.5, [[("Qdrant ", {"bold": True}), ("stores one point per chunk: the vector to search by, and the "
                               "payload you get back, with the text, the file and the page.", {})]], size=18, line=1.25)
notes(s, "3 minutes. The cat/kitten numbers are the real ones participants will see in notebook 3.")

s = notebook(3, "Search", 13,
             ["Embed sentences and compare how similar they are", "Store all chunks in Qdrant, in memory",
              "Turn the bot into a search engine, no LLM yet"],
             ["Fewer, better results: pick a score cutoff", "Search one paper only: a dropdown and a filter",
              "If time: open each hit as a PDF page"],
             [("you", "What were the lifetime-encoded beads loaded with?", 0.62),
              ("bot", "Kage_2018_SciReports.pdf, p. 2  (score 0.73)\n› … loaded with different organic fluorophores …\n\n"
                      "Kage_2018_SciReports.pdf, p. 3  (score 0.64)\n› Table 1. Lifetime codes and respective luminophores …", 2.55),
              ("you", "What is the capital of France?", 0.62),
              ("bot", "Kage_2018_SciReports.pdf, p. 6  (score 0.17) …", 0.75)])
notes(s, "13 minutes. Make them notice that search always returns something, even for France: that's why "
         "exercise 1 asks for a score cutoff. Good questions scored 0.33 to 0.73 in testing, France 0.17.")

s = content("Prompt and citations", BLUE, "BASICS")
diagram("rag-prompt", 1.16, 2.0, 11.4, s)
diagram("citations", 1.16, 6.5, 9.3, s)
for i, (word, rest) in enumerate([("Retrieve", "the 5 closest chunks"), ("Augment", "the prompt with them, numbered [1] to [5]"),
                                  ("Generate", "the answer, citing those numbers")]):
    text(s, 13.4, 2.3 + i * 1.25, 6.2, 1.1, [[(word.upper(), {"bold": True, "color": BLUE, "spacing": 3, "size": 17})],
                                             [(rest, {"size": 19})]])
text(s, 13.4, 6.9, 6.0, 2.5, "Every [n] leads back to a chunk, and every chunk remembers its file and page. "
     "That is all a citation is.", size=19, line=1.25)
notes(s, "3 minutes. RAG = retrieve, augment, generate. Citations are bookkeeping, not magic.")

s = notebook(4, "Retrieval-augmented generation", 13,
             ["Ask notebook 1's question again, this time with context", "Let the model cite the numbered chunks",
              "Turn every [n] into a link to the PDF page"],
             ["Strict or not: remove \"say you don't know\"", "How much context: limit 2 versus 15",
              "Show the evidence in the search step", "If time: follow-up questions"],
             [("you", "In Kage et al. (2018), what were the lifetime-encoded beads loaded with?", 1.0),
              ("step", "Search the papers", 0.45),
              ("bot", "PMMA beads stained with organic dyes from PolyAn GmbH [1], and melamine beads loaded with "
                      "CdSe/CdS/ZnS quantum dots [1], [2].\n\nSources:  Kage_2018_SciReports, p. 2", 2.4),
              ("step", "Click the source: the PDF opens at page 2.", 0.45)])
notes(s, "13 minutes. Compare with the invented answer from notebook 1, same question. Let them click a source.")

s = content("Putting it together", YELLOW, "HANDS-ON")
kicker(s, 1.16, 1.85, 9, "Notebook 5 · 10 min", ORANGE)
diagram("overview", 1.16, 2.45, 10.3, s)
text(s, 1.16, 6.35, 10.2, 1.0, "One app.py with every piece. Attach a PDF in the chat and the top row runs live: "
     "read, chunk, embed, store. Then ask about it.", size=18, line=1.25)
kicker(s, 12.3, 2.45, 7, "Your turn · in the chatbot")
numbered(s, 12.3, 2.95, 7.2, ["Your own document: attach a PDF and ask", "Chat with this document only: a Qdrant filter",
                              "If time: a dropdown to switch the model"], gap=0.95)
text(s, 12.3, 5.85, 7.2, 0.4, "Solutions are folded away under each exercise.", size=15, color=MUTED)
pieces = [("Read a PDF with Docling", "2"), ("Cut into chunks", "2"), ("Turn text into vectors", "3"),
          ("Store and search in Qdrant", "3"), ("Answer from the chunks", "4"), ("Cite the page", "4")]
kicker(s, 1.16, 7.7, 9, "Piece · notebook")
for i, (piece, nb) in enumerate(pieces):
    x, y = 1.16 + (i % 3) * 6.0, 8.2 + (i // 3) * 0.95
    rect(s, x, y, 5.7, 0.75, SAND)
    text(s, x + 0.25, y, 4.7, 0.75, piece, size=17, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + 4.9, y, 0.6, 0.75, nb, size=20, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
notes(s, "10 minutes. Remind them how the pieces fit: everything from notebooks 1 to 4, one file. "
         "Only test documents, nothing confidential, when they attach their own PDF.")

# ---------------------------------------------------------------------------- 24–25 wrap-up
s = content("Where to go next", BLUE, "NEXT STEPS")
rows = [("Today", "The template"), ("PDFs only", "Also txt, md, csv and json"),
        ("Fixed 1000-character chunks", "Structure-aware chunkers, chosen per data source in a config file"),
        ("Qdrant in memory, rebuilt on every start", "Qdrant server, ingested once, updated when files change"),
        ("Vector search only", "Hybrid search: vectors plus keyword matching"),
        ("[1] and a file name", "Citation format set in the config: title, file, page"),
        ("You judge the answers by eye", "An evaluation app that scores answers")]
for i, (today, template) in enumerate(rows):
    y = 2.15 + i * 0.86
    fill = BLUE if i == 0 else ("F2F1EE" if i % 2 else WHITE)
    rect(s, 1.16, y, 17.7, 0.82, fill)
    ink, b = (WHITE, True) if i == 0 else (INK, False)
    text(s, 1.45, y, 6.5, 0.82, today, size=18, bold=b, color=ink, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 8.3, y, 10.4, 0.82, template, size=18, bold=b, color=ink, anchor=MSO_ANCHOR.MIDDLE)
for i, (label, value) in enumerate([("This workshop", REPO), ("The template", "github.com/aihpi/pilotproject-rag-template"),
                                    ("Other embedding models", "huggingface.co/spaces/mteb/leaderboard")]):
    kicker(s, 1.16 + i * 6.0, 8.6, 5.8, label)
    text(s, 1.16 + i * 6.0, 9.05, 5.9, 0.8, [[(value, {"link": "https://" + value})]], size=15, font="Courier New", color="DC640D")
notes(s, "2 minutes. The template is where to start a real project, so nobody starts from zero.")

s = prs.slides.add_slide(BLANK)
text(s, 1.67, 1.85, 13.5, 1.2, "YOUR FEEDBACK COUNTS", size=60, bold=True, color=GREY)
text(s, 1.67, 3.1, 9.5, 1.2, "We shape our next workshops around your answers. Two minutes is enough.", size=24, line=1.25)
rect(s, 1.67, 4.55, 1.25, 0.08, RED)
image(s, "qr-feedback.png", 1.67, 4.97, w=2.92)
text(s, 5.05, 5.6, 6, 2.2, [[("kisz@hpi.de", {"link": "mailto:kisz@hpi.de"})], [("hpi.de/kisz", {"link": "https://hpi.de/kisz"})],
                            [("[Presenter]", {"color": GREY, "size": 17})]], size=22, line=1.4)
rect(s, 1.67, 8.3, 7.4, 0.02, "D0D3D6")
text(s, 1.67, 8.5, 8, 0.4, "KEEP GOING", size=15, bold=True, color=RED, spacing=3)
text(s, 1.67, 8.95, 8.5, 1.0, [[(REPO, {"link": "https://" + REPO}), ("  ·  this workshop", {"color": GREY})],
                               [("github.com/aihpi/pilotproject-rag-template", {"link": "https://github.com/aihpi/pilotproject-rag-template"}),
                                ("  ·  template", {"color": GREY})]], size=15, line=1.3)
image(s, "logo-kisz.png", 16.45, 0.74, w=2.8)
image(s, "logo-bmftr.png", 16.41, 1.98, w=2.8)
image(s, "hpi-tagline.png", 0.76, 10.23, w=2.92)
image(s, "hpi-building.png", 9.72, 6.72, w=11.72)
page_number(s)
notes(s, "Leave the QR code up while questions come in. Contact: kisz@hpi.de.")

prs.save(HERE / "rag-workshop.pptx")
print("wrote rag-workshop.pptx,", len(prs.slides), "slides")
