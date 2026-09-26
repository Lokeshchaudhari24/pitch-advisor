from src.retrieval import search_policy

query = "Which policy provides unlimited automatic recharge?"

results = search_policy(query)

for i, document in enumerate(results["documents"][0]):
    print(f"\n--- Result {i+1} ---")
    print(document)

    print("Source:")
    print(results["metadatas"][0][i])