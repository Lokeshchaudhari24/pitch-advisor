from src.ingestion import extract_pdf
from src.chunking import chunk_documents

documents = extract_pdf(
    "data/Care Health Product Brochure.pdf"
)

chunks = chunk_documents(documents)

print("Pages:", len(documents))
print("Chunks:", len(chunks))

print("\nFirst chunk:")
print(chunks[0]["text"])

print("\nMetadata:")
print(chunks[0]["metadata"])