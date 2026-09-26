import pymupdf
from pathlib import Path


def extract_pdf(pdf_path):
    doc = pymupdf.open(pdf_path)

    documents = []

    for page_number, page in enumerate(doc):
        text = page.get_text("text")

        if text.strip():
            documents.append({
                "text": text,
                "metadata": {
                    "source": Path(pdf_path).name,
                    "page": page_number + 1
                }
            })

    return documents