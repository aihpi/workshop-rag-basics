"""Writes the workshop diagrams to workshop/img/. Run from this folder: python make_svgs.py ."""
import sys
from pathlib import Path
from xml.sax.saxutils import escape

OUT = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
KIND = {  # fill, stroke
    "data": ("#e8f0fe", "#3b6fd8"),
    "step": ("#fff1e0", "#d9822b"),
    "model": ("#f1ebff", "#7c5cd6"),
    "store": ("#e6f6ec", "#2f9e62"),
    "plain": ("#ffffff", "#c5ccd3"),
}
INK, MUTED = "#1f2933", "#5f6b7a"
STYLE = f"""<style>
text {{ font-family: -apple-system, "Segoe UI", Helvetica, Arial, sans-serif; fill: {INK}; }}
.title {{ font-size: 14px; font-weight: 600; }}
.sub {{ font-size: 12px; fill: {MUTED}; }}
.lane {{ font-size: 11px; font-weight: 700; letter-spacing: .08em; fill: {MUTED}; }}
.label {{ font-size: 12px; fill: {MUTED}; }}
.mono {{ font-family: Menlo, Consolas, "DejaVu Sans Mono", monospace; font-size: 12px; }}
</style>
<defs><marker id="arrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
<path d="M0,0 L10,5 L0,10 z" fill="{MUTED}"/></marker></defs>"""


def svg(name, w, h, body, alt):
    OUT.joinpath(name).write_text(
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w} {h}" width="{w}" role="img" aria-label="{escape(alt)}">\n'
        f"<title>{escape(alt)}</title>\n{STYLE}\n"
        f'<rect x="1" y="1" width="{w - 2}" height="{h - 2}" rx="14" fill="#ffffff" stroke="#e1e5ea"/>\n'
        + "\n".join(body) + "\n</svg>\n"
    )


def box(x, y, w, h, kind, title=None, sub=None, rx=10):
    fill, stroke = KIND[kind]
    out = [f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" fill="{fill}" stroke="{stroke}" stroke-width="1.5"/>']
    cx = x + w / 2
    if title and sub:
        out.append(f'<text class="title" x="{cx}" y="{y + h / 2 - 4}" text-anchor="middle">{escape(title)}</text>')
        out.append(f'<text class="sub" x="{cx}" y="{y + h / 2 + 14}" text-anchor="middle">{escape(sub)}</text>')
    elif title:
        out.append(f'<text class="title" x="{cx}" y="{y + h / 2 + 5}" text-anchor="middle">{escape(title)}</text>')
    return out


def arrow(points, label=None, lx=None, ly=None, dashed=False):
    d = "M" + " L".join(f"{x},{y}" for x, y in points)
    dash = ' stroke-dasharray="5 4"' if dashed else ""
    out = [f'<path d="{d}" fill="none" stroke="{MUTED}" stroke-width="1.5"{dash} marker-end="url(#arrow)"/>']
    if label:
        (x1, y1), (x2, y2) = points[0], points[-1]
        lx = (x1 + x2) / 2 if lx is None else lx
        ly = (y1 + y2) / 2 - 8 if ly is None else ly
        out.append(f'<text class="label" x="{lx}" y="{ly}" text-anchor="middle">{escape(label)}</text>')
    return out


def text(x, y, s, cls="label", anchor="start", extra=""):
    if cls == "mono":
        extra += ' xml:space="preserve"'  # keep the padding that lines up table columns
    return [f'<text class="{cls}" x="{x}" y="{y}" text-anchor="{anchor}"{extra}>{escape(s)}</text>']


# 1 · overview ---------------------------------------------------------------
W, GAP, X0 = 136, 22, 24
col = lambda i: X0 + i * (W + GAP)
ya, yb, h = 66, 226, 58
b = text(24, 48, "ONCE, AHEAD OF TIME", "lane") + text(24, 208, "FOR EVERY QUESTION", "lane")
for i, (k, t, s) in enumerate([
    ("data", "PDFs", "3 papers"),
    ("step", "Docling", "PDF → Markdown"),
    ("data", "Chunks", "≈ 1000 characters"),
    ("model", "Embedding model", "text → 4096 numbers"),
    ("store", "Qdrant", "vectors + text + page"),
]):
    b += box(col(i), ya, W, h, k, t, s)
    if i:
        b += arrow([(col(i) - GAP, ya + h / 2), (col(i) - 2, ya + h / 2)])
for i, (k, t, s) in enumerate([
    ("data", "Question", "typed in the chat"),
    ("model", "Embedding model", "same model"),
    ("store", "Search", "5 closest chunks"),
    ("step", "Prompt", "chunks + question"),
    ("model", "LLM", "gemma-4-31b"),
    ("data", "Answer", "with [1] [2] links"),
]):
    b += box(col(i), yb, W, h, k, t, s)
    if i:
        b += arrow([(col(i) - GAP, yb + h / 2), (col(i) - 2, yb + h / 2)])
qx, sx = col(4) + W / 2, col(2) + W / 2
b += arrow([(qx, ya + h), (qx, 176), (sx, 176), (sx, yb - 2)], dashed=True)
b += text((qx + sx) / 2, 168, "searches the same collection", anchor="middle")
lx = 24
for k, name in [("data", "data"), ("step", "processing step"), ("model", "AI model"), ("store", "vector database")]:
    fill, stroke = KIND[k]
    b.append(f'<rect x="{lx}" y="312" width="14" height="14" rx="3" fill="{fill}" stroke="{stroke}" stroke-width="1.5"/>')
    b += text(lx + 20, 324, name)
    lx += 40 + len(name) * 7
svg("overview.svg", col(5) + W + 24, 344, b,
    "The two halves of RAG. Ahead of time: PDFs go through Docling, are cut into chunks, embedded and stored in Qdrant. "
    "For every question: the question is embedded, the 5 closest chunks are found in Qdrant, put into a prompt with the question, "
    "and the LLM writes an answer with source links.")

# 2 · live reload --------------------------------------------------------------
b = []
b += box(24, 40, 200, 64, "step", "Notebook cell", "%%writefile app.py")
b += box(290, 40, 150, 64, "data", "app.py", "the bot's code")
b += box(506, 40, 180, 64, "model", "Chainlit server", "watches app.py")
b += box(752, 40, 150, 64, "plain", "Browser", "localhost:8000")
b += arrow([(226, 72), (288, 72)], "writes")
b += arrow([(442, 72), (504, 72)], "notices")
b += arrow([(688, 72), (750, 72)], "reloads")
b += text(463, 140, "Run the cell, switch to the browser tab: the bot has a new skill.", anchor="middle")
svg("live-reload.svg", 926, 164, b,
    "A notebook cell writes app.py, the Chainlit server notices the change and reloads the chat in the browser.")

# 3 · docling ----------------------------------------------------------------
b = text(40, 30, "PDF PAGE", "lane") + text(470, 30, "MARKDOWN FROM DOCLING", "lane")
b.append('<rect x="40" y="42" width="220" height="282" rx="4" fill="#ffffff" stroke="#c5ccd3" stroke-width="1.5"/>')
regions = [  # y, h, kind, label
    (56, 20, "data", "heading"),
    (84, 52, "plain", "paragraph"),
    (144, 92, "store", "table"),
    (244, 66, "model", "figure"),
]
for y, hh, k, label in regions:
    fill, stroke = KIND[k]
    b.append(f'<rect x="54" y="{y}" width="192" height="{hh}" rx="3" fill="{fill}" stroke="{stroke}" stroke-width="1.5" stroke-dasharray="4 3"/>')
    b += text(240, y + 14, label, "label", "end")
for y in (108, 118, 128):  # paragraph lines
    b.append(f'<line x1="62" y1="{y}" x2="196" y2="{y}" stroke="#c5ccd3" stroke-width="3" stroke-linecap="round"/>')
for y in (170, 190, 210):  # table rows
    b.append(f'<line x1="62" y1="{y}" x2="238" y2="{y}" stroke="#2f9e62" stroke-width="1"/>')
for x in (120, 180):
    b.append(f'<line x1="{x}" y1="160" x2="{x}" y2="228" stroke="#2f9e62" stroke-width="1"/>')
b += arrow([(272, 183), (452, 183)])
b += text(362, 165, "finds headings, tables,", anchor="middle") + text(362, 205, "figures, reading order", anchor="middle")
md = [
    (66, "data", "## Results and discussion"),
    (100, "plain", "Our experimental approach presented"),
    (118, "plain", "in Fig. 1 required dye/size-encoded …"),
    (160, "store", "| Aptamer    | Target   | Buffer  |"),
    (178, "store", "|------------|----------|---------|"),
    (196, "store", "| T1-apta 17 | Thrombin | 20 mM … |"),
    (214, "store", "| T2-apta 18 | Thrombin | 50 mM … |"),
    (266, "model", "Figure 1. Principle of the multiplexed …"),
    (284, "model", "<!-- image -->"),
]
for y, k, line in md:
    b.append(f'<rect x="470" y="{y - 13}" width="4" height="18" fill="{KIND[k][1]}"/>')
    b += text(484, y, line, "mono")
svg("docling.svg", 820, 344, b,
    "Docling recognises the parts of a PDF page (heading, paragraph, table, figure) and turns them into Markdown, "
    "keeping the heading as ## and the table as a Markdown table.")

# 4 · chunking ---------------------------------------------------------------
scale, x0 = 0.3, 40  # px per character
X = lambda c: x0 + c * scale
b = text(x0, 36, "ONE PAGE OF TEXT · 2,600 CHARACTERS", "lane")
for a, z in [(800, 1000), (1600, 1800)]:
    b.append(f'<rect x="{X(a)}" y="48" width="{(z - a) * scale}" height="164" fill="#d9822b" opacity=".14"/>')
    b += text((X(a) + X(z)) / 2, 236, "200 shared", anchor="middle")
b.append(f'<rect x="{X(0)}" y="52" width="{2600 * scale}" height="26" rx="4" fill="#eef1f4" stroke="#c5ccd3"/>')
for i, (a, z) in enumerate([(0, 1000), (800, 1800), (1600, 2600)]):
    y = 96 + i * 38
    b += box(X(a), y, (z - a) * scale, 28, "data", rx=5)
    b += text((X(a) + X(z)) / 2, y + 19, f"chunk {i + 1} · characters {a}–{z}", "label", "middle", ' style="fill:#1f2933"')
b += arrow([(X(0), 262), (X(800) - 2, 262)])
b += text((X(0) + X(800)) / 2, 280, "next chunk starts 800 later (size 1000 − overlap 200)", anchor="middle")
svg("chunking.svg", 2600 * scale + 2 * x0, 300, b,
    "A page of 2600 characters is cut into three chunks of 1000 characters. Each chunk starts 800 characters after the previous one, "
    "so neighbouring chunks share 200 characters.")

# 5 · embeddings ---------------------------------------------------------------
import math

ox, oy = 70, 300
b = [f'<line x1="{ox}" y1="{oy}" x2="{ox + 330}" y2="{oy}" stroke="#c5ccd3" stroke-width="1.5"/>',
     f'<line x1="{ox}" y1="{oy}" x2="{ox}" y2="{oy - 250}" stroke="#c5ccd3" stroke-width="1.5"/>']
pt = lambda deg, r: (ox + r * math.cos(math.radians(deg)), oy - r * math.sin(math.radians(deg)))
points = [(58, "The cat sits on the mat.", "#3b6fd8"), (46, "A kitten is resting on a rug.", "#3b6fd8"),
          (9, "Quarterly revenue grew by four percent.", "#d9822b")]
for deg, label, colour in points:
    x, y = pt(deg, 230)
    b.append(f'<line x1="{ox}" y1="{oy}" x2="{x:.1f}" y2="{y:.1f}" stroke="{colour}" stroke-width="2"/>')
    b.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="5" fill="{colour}"/>')
    b += text(x + 10, y + 4, f"“{label}”", "label", extra=' style="fill:#1f2933"')


def arc(a, z, r, label):
    (x1, y1), (x2, y2) = pt(a, r), pt(z, r)
    lx, ly = pt(z + 20, r) if z - a < 20 else pt((a + z) / 2, r + 22)  # narrow wedge: label beside it
    return [f'<path d="M{x1:.1f},{y1:.1f} A{r},{r} 0 0,0 {x2:.1f},{y2:.1f}" fill="none" stroke="{MUTED}" stroke-width="1.2"/>'] + \
        text(lx, ly + 4, label, "title", "middle")


b += arc(46, 58, 90, "0.57") + arc(9, 58, 150, "0.13")
b += text(ox, 30, "Similar meaning → arrows point the same way.", "title")
b += text(ox, 330, "Similarity = how closely two arrows point the same way (1 = same direction).", "label")
b += text(ox, 348, "Real embeddings have 4096 dimensions; this drawing has 2.", "label")
svg("embeddings.svg", 620, 368, b,
    "Embeddings drawn as arrows. 'The cat sits on the mat' and 'A kitten is resting on a rug' point in almost the same direction "
    "(similarity 0.57); 'Quarterly revenue grew by four percent' points elsewhere (0.13).")

# 6 · qdrant -----------------------------------------------------------------
b = []
b += box(24, 24, 540, 258, "store", rx=12)
b += text(42, 50, "Collection “papers”", "title")
b += text(42, 80, "id", "lane") + text(96, 80, "VECTOR", "lane") + text(316, 80, "PAYLOAD", "lane")
rows = [("0", "[0.012, -0.031, 0.087, …]", "Kage_2018 · p. 1 · “Time-res…”"),
        ("1", "[-0.007, 0.044, 0.019, …]", "Kage_2018 · p. 1 · “multiplex…”"),
        ("2", "[0.021, 0.003, -0.058, …]", "Kage_2018 · p. 2 · “are loade…”")]
for i, (pid, vec, payload) in enumerate(rows):
    y = 92 + i * 52
    b.append(f'<rect x="38" y="{y}" width="512" height="42" rx="6" fill="#ffffff" stroke="#2f9e62" stroke-width="1"/>')
    b += text(48, y + 26, pid, "mono") + text(96, y + 26, vec, "mono") + text(316, y + 26, payload, "mono")
b += text(294, 266, "… 203 points, one per chunk", "label", "middle")
b += box(640, 40, 196, 60, "model", "Question → vector", "embedded like the chunks")
b += box(640, 206, 196, 60, "data", "5 closest points", "each with score + payload")
b += arrow([(640, 70), (566, 70)], "query_points", 603, 60)
b += arrow([(566, 236), (638, 236)], "returns", 603, 226)
b += text(96, 316, "vector = what Qdrant searches by", "label") + text(316, 316, "payload = what you get back", "label")
svg("qdrant.svg", 860, 336, b,
    "A Qdrant collection called papers holds one point per chunk. Each point has an id, a vector that Qdrant searches by, "
    "and a payload with file, page and text that comes back with the results. A question vector goes in, the 5 closest points come out.")

# 7 · rag prompt ---------------------------------------------------------------
b = text(24, 34, "WHAT THE LLM RECEIVES", "lane")
b += box(24, 46, 500, 64, "step", rx=8)
b += text(38, 68, "system", "lane") + text(38, 90, "Answer only from the context. If it isn't there, say you don't know.", "label",
                                          extra=' style="fill:#1f2933"')
b += box(24, 122, 500, 178, "plain", rx=8)
b += text(38, 144, "user", "lane") + text(38, 166, "Context:", "mono")
for y, line in [(188, "[1] …microbeads are loaded with different organic…"),
                (212, "[2] …tailor-made melamine beads with a polyelectr…"),
                (256, "[5] …imitations of lifetime encoding in flow with…")]:
    b.append(f'<rect x="38" y="{y - 15}" width="472" height="20" rx="4" fill="{KIND["store"][0]}"/>')
    b += text(46, y, line, "mono")
b += text(46, 236, "⋮", "mono")
b.append(f'<rect x="38" y="272" width="3" height="20" fill="{KIND["data"][1]}"/>')
b += text(46, 287, "Question: what were the beads loaded with?", "mono")
b += arrow([(526, 173), (582, 173)])
b += box(584, 143, 100, 60, "model", "LLM", "gemma-4-31b")
b += arrow([(686, 173), (726, 173)])
b += box(728, 110, 168, 126, "data", rx=8)
for i, line in enumerate(["The beads were loaded", "with different organic", "fluorophores with", "lifetimes in the", "nanosecond range [1]."]):
    b += text(740, 136 + i * 20, line, "label", extra=' style="fill:#1f2933"')
b += text(274, 322, "green: the 5 chunks Qdrant found · blue: what you typed", "label", "middle")
svg("rag-prompt.svg", 920, 340, b,
    "The prompt sent to the LLM: a system message with the rules, then a user message with the 5 retrieved chunks numbered [1] to [5] "
    "and the question. The LLM answers from those chunks and cites [1].")

# 8 · citations ----------------------------------------------------------------
b = []
b += box(24, 56, 210, 96, "data", rx=8)
for i, line in enumerate(["…loaded with different", "organic fluorophores", "in the nanosecond", "range"]):
    b += text(38, 82 + i * 19, line, "label", extra=' style="fill:#1f2933"')
b.append(f'<rect x="80" y="125" width="26" height="19" rx="4" fill="#ffd9a8"/>')
b += text(93, 139, "[1]", "mono", "middle")
b += arrow([(236, 104), (300, 104)], "finds [1]", 268, 94)
b += box(302, 56, 250, 96, "store", rx=8)
b += text(316, 82, "hits[0].payload", "mono")
b += text(316, 106, "source: Kage_2018_…pdf", "mono") + text(316, 126, "page:   2", "mono")
b += arrow([(554, 104), (618, 104)], "cl.Pdf", 586, 94)
b.append('<path d="M620,30 h130 l24,24 v130 h-154 z" fill="#ffffff" stroke="#c5ccd3" stroke-width="1.5"/>')
b.append('<path d="M750,30 v24 h24" fill="none" stroke="#c5ccd3" stroke-width="1.5"/>')
for y in (74, 86, 98, 110, 122, 134):
    b.append(f'<line x1="634" y1="{y}" x2="{748 if y % 24 else 722}" y2="{y}" stroke="#c5ccd3" stroke-width="3" stroke-linecap="round"/>')
b += text(697, 168, "opens at page 2", "label", "middle")
b += text(24, 210, "The number in the answer points into the list of chunks, and every chunk remembers its file and page.", "label")
svg("citations.svg", 800, 232, b,
    "The answer cites [1]. [1] is the first search hit, whose payload says Kage_2018_SciReports.pdf, page 2. "
    "Chainlit's cl.Pdf element opens that file at page 2.")
print("wrote", sorted(p.name for p in OUT.glob("*.svg")))
