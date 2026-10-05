# Workshop: Build a RAG chatbot

You build a chatbot that answers questions about three scientific papers and links every
answer to the page it comes from. Five short notebooks build it piece by piece, and you watch
it change in your browser. Each notebook ends with small exercises that you do in the chatbot.

## Before the workshop

You need [uv](https://docs.astral.sh/uv/getting-started/installation/) (it installs Python
for you) and Git.

```bash
git clone https://github.com/aihpi/workshop-rag-basics.git
cd workshop-rag-basics/workshop
uv sync
cp .env.example .env
```

Open `.env` and put in the key you got from the workshop host. Then check everything works:

```bash
uv run python check_setup.py
```

This also downloads the models [Docling](https://github.com/docling-project/docling) uses
to read PDFs. Altogether you need about 2 GB of disk space, so please do this at home rather
than over the workshop Wi-Fi.

## On the day

```bash
uv run jupyter lab
```

Open the notebooks in order. The chatbot runs at http://localhost:8000.

| Time | | |
|---|---|---|
| 0:00 | Theory | What RAG is, and why |
| 0:15 | `01_chat.ipynb` | Talking to an LLM; your first Chainlit app |
| 0:30 | `02_documents.ipynb` | Reading PDFs with Docling; chunks |
| 0:45 | Theory | Embeddings and vector search |
| 0:55 | `03_search.ipynb` | Embeddings; Qdrant; a search engine |
| 1:15 | `04_rag.ipynb` | Answering from the papers, with citations |
| 1:35 | `05_together.ipynb` | All pieces in one chatbot; your own PDFs |
| 1:45 | End | |

Each notebook needs the files the earlier ones save (`chunks.json`, `index.json`), so run
them in order. The solutions to the exercises are folded away under each exercise list.

The papers in `data/` are open access under CC BY 4.0; see [data/SOURCES.md](data/SOURCES.md).
