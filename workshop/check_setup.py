"""Run once before the workshop: uv run python check_setup.py

Checks your gateway key and downloads Docling's models, so nothing big has to
download over the workshop Wi-Fi.
"""
import os

from docling.datamodel.base_models import InputFormat
from docling.datamodel.pipeline_options import PdfPipelineOptions
from docling.document_converter import DocumentConverter, PdfFormatOption
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()
llm = OpenAI(base_url=os.environ["LLM_BASE_URL"], api_key=os.environ["LLM_API_KEY"])

print("Chat model ...", end=" ", flush=True)
llm.chat.completions.create(model=os.environ["CHAT_MODEL"], messages=[{"role": "user", "content": "Hi"}], max_tokens=5)
print("ok")

print("Embedding model ...", end=" ", flush=True)
llm.embeddings.create(model=os.environ["EMBED_MODEL"], input=["Hi"])
print("ok")

print("Docling (the first run downloads its models, a few minutes) ...", flush=True)
converter = DocumentConverter(
    format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=PdfPipelineOptions(do_ocr=False))}
)
converter.convert("data/Schmidt_2022_SciReports.pdf")
print("ok\n\nYou're ready for the workshop.")
