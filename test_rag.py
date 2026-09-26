from src.retrieval import search_policy
from src.generation import generate_answer

question = "Which policies provide automatic recharge or restore benefits?"

results = search_policy(question, top_k=5)

print("🔍 --- CHROMADB NE YEH CHUNKS NIKALE --- 🔍")

# YAHAN CHANGE HAI: Ab hum har chunk ke upar uski PDF ka naam bhi likh rahe hain
context_chunks = []
for idx, doc in enumerate(results["documents"][0]):
    source_pdf = results["metadatas"][0][idx]["source"]
    chunk_with_source = f"--- Source Document: {source_pdf} ---\n{doc}"
    context_chunks.append(chunk_with_source)

context = "\n\n".join(context_chunks)
print(context)
print("------------------------------------------\n")

answer = generate_answer(question, context)
print("🤖 --- AI KA FINAL ANSWER --- 🤖")
print(answer)