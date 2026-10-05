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
(it installs Python for you) and Git. Bring a test or public PDF of your own for notebook 5.

## On the day

Start the install as soon as you sit down; it downloads about 2 GB:

```bash
git clone https://github.com/aihpi/workshop-rag-basics.git
cd workshop-rag-basics/workshop
uv sync
```

The gateway key is handed out on the day. Put it into `.env`, then check everything works and
fetch the models [Docling](https://github.com/docling-project/docling) uses to read PDFs:

```bash
cp .env.example .env    # then paste the key
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
| 0:00 | Welcome, route, setup | The install runs while we talk |
| 0:15 | `01_chat.ipynb` | Talking to an LLM: the confident answer |
| 0:28 | `02_documents.ipynb` | Reading PDFs with Docling; chunks |
| 0:41 | `03_search.ipynb` | Embeddings, Qdrant, a search engine |
| 0:54 | `04_rag.ipynb` | Answers with citations; which claims are backed |
| 1:10 | `05_together.ipynb` | Your own PDF; what would break on your collection; next steps |
| 1:30 | End | |

The slides are in [`slides/`](slides/): [`rag-workshop.pdf`](slides/rag-workshop.pdf) to present,
[`rag-workshop.pptx`](slides/rag-workshop.pptx) with speaker notes. Everything after the
appendix divider is background for questions.

The papers in `data/` are open access under CC BY 4.0; see [data/SOURCES.md](data/SOURCES.md).
