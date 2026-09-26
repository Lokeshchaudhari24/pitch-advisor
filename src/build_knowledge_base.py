from pathlib import Path

from src.ingestion import extract_pdf
from src.chunking import chunk_documents
from src.embeddings import create_embeddings
from src.vector_store import store_chunks


data_folder = Path("data")

all_documents = []

for pdf_path in data_folder.glob("*.pdf"):

    print(f"Processing: {pdf_path.name}")

    documents = extract_pdf(str(pdf_path))
    all_documents.extend(documents)


chunks = chunk_documents(all_documents)

embeddings = create_embeddings(chunks)

store_chunks(chunks, embeddings)

print(f"\nTotal PDFs: {len(list(data_folder.glob('*.pdf')))}")
print(f"Total pages: {len(all_documents)}")
print(f"Total chunks: {len(chunks)}")