# Workshop: Build a RAG chatbot in 90 minutes

You build a chatbot that answers questions about three scientific papers and links every
answer to the page it comes from. You write it cell by cell in a Jupyter notebook and
watch it grow in your browser.

## Before the workshop

You need [uv](https://docs.astral.sh/uv/getting-started/installation/) (it installs Python
for you) and Git.

```bash
git clone <this repository>
cd <repository>/workshop
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
uv run jupyter lab workshop.ipynb
```

Work through the notebook from top to bottom. The chatbot runs at http://localhost:8000.

The papers in `data/` are open access under CC BY 4.0; see [data/SOURCES.md](data/SOURCES.md).
