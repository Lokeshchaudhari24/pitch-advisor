from src.ingestion import extract_pdf

documents = extract_pdf(
    "data/Care Health Product Brochure.pdf"
)

print("Pages extracted:", len(documents))

print("\nFirst page:")
print(documents[0]["text"])

print("\nMetadata:")
print(documents[0]["metadata"])