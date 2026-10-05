"""Run once you have the key: uv run python check_setup.py

Checks your gateway key and the two models, and downloads Docling's models
(about 500 MB) so notebook 2 doesn't have to wait for them.
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
