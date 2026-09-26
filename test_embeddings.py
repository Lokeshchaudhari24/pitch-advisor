from src.ingestion import extract_pdf
from src.chunking import chunk_documents
from src.embeddings import create_embeddings

documents = extract_pdf(
    "data/Care Health Product Brochure.pdf"
)

chunks = chunk_documents(documents)

embeddings = create_embeddings(chunks)

print("Number of chunks:", len(chunks))
print("Embedding shape:", embeddings.shape)
print("First embedding:")
print(embeddings[0])