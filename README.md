# Hands-On RAG: Build Your Own Document Q&A System

You build a chatbot that answers questions about three scientific papers, and then about a PDF of
your own, and links every claim to the page it comes from, with a score that says how well the
answer is backed. Five short notebooks build it piece by piece, and you watch it change in your
browser.

**Who this is for:** researchers and practitioners who can read and edit a short Python function.
No knowledge of RAG or embeddings needed.

## Before the workshop

Work through the setup guide at [aihpi/workshop-getting-started](https://github.com/aihpi/workshop-getting-started/).
Of what it installs, this workshop needs [uv](https://docs.astral.sh/uv/getting-started/installation/)
(it installs Python for you) and Git. You also need about 2.5 GB of free disk space.

Bring a PDF of your own for notebook 5: test or public documents only, with real text (not a
scan), and short, ideally under 20 pages. No PDF of your own? Any open-access paper works.

If you like, run the install below at home already: it doesn't need the key and saves the 2 GB
download on the workshop Wi-Fi.

## On the day

Start the install as soon as you sit down; it downloads about 2 GB:

```bash
git clone https://github.com/aihpi/workshop-rag-basics.git
cd workshop-rag-basics
uv sync
```

The gateway key is handed out on the day. Put it into `.env`, then check everything works and
fetch the models [Docling](https://github.com/docling-project/docling) uses to read PDFs:

```bash
cp .env.example .env
open -e .env            # Windows: notepad .env; paste the key after LLM_API_KEY=
uv run python check_setup.py
uv run jupyter lab
```

Open the notebooks in order; each needs the files the one before saves (`chunks.json`,
`index.json`). The chatbot runs at http://localhost:8000. Solutions are folded away under each
exercise.

**Your data:** everything you send, including your own PDF, goes to the AISC gateway on our
infrastructure in Potsdam. Use test or public documents only.

## Agenda

| Time | Session | |
|---|---|---|
| 0:00 | Welcome, the AI Service Centre, setup | The install runs while we talk |
| 0:22 | `01_chat.ipynb` | Talking to an LLM: the confident answer |
| 0:35 | `02_documents.ipynb` | Reading PDFs with Docling; chunks |
| 0:48 | `03_search.ipynb` | Embeddings, Qdrant, a search engine |
| 1:01 | Break | |
| 1:06 | `04_rag.ipynb` | Answers with citations; which claims are backed |
| 1:22 | `05_together.ipynb` | Your own PDF; what would go wrong with your documents |
| 1:40 | Questions, feedback | |
| 1:45 | End | |

Every notebook has a fixed time. Exercises marked **Bonus** are only for when you finish early;
when the time is up, we move on together.

The slides are in [`slides/`](slides/): [`rag-workshop.pdf`](slides/rag-workshop.pdf) to present,
[`rag-workshop.pptx`](slides/rag-workshop.pptx) with speaker notes. Everything after the
appendix divider is background for questions.

## After the workshop

This repository is the workshop: small on purpose, so you can read every line. For a real project,
start from the [RAG template](https://github.com/aihpi/pilotproject-rag-template). It runs the same
steps with more file types, structure-aware chunking, hybrid search, a persistent Qdrant server and
an evaluation app, all set in one config file.

The papers in `data/` are open access under CC BY 4.0; see [data/SOURCES.md](data/SOURCES.md).
