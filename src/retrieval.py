from src.vector_store import collection
from src.embeddings import get_embedding_model


def search_policy(query, top_k=5):
    model = get_embedding_model()

    query_embedding = model.encode(
        [query],
        normalize_embeddings=True
    )

    results = collection.query(
        query_embeddings=query_embedding.tolist(),
        n_results=top_k
    )

    return results