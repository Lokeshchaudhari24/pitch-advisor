from src.ingestion import extract_pdf
from src.chunking import chunk_documents
from src.embeddings import create_embeddings
from src.vector_store import store_chunks


documents = extract_pdf(
    "data/Care Health Product Brochure.pdf"
)

chunks = chunk_documents(documents)

embeddings = create_embeddings(chunks)

store_chunks(chunks, embeddings)