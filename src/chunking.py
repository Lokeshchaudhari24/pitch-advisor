def chunk_documents(documents, chunk_size=500, overlap=100):

    chunks = []

    for doc in documents:
        text = doc["text"]

        start = 0

        while start < len(text):

            end = start + chunk_size

            chunks.append({
                "text": text[start:end],
                "metadata": doc["metadata"]
            })

            start += chunk_size - overlap

    return chunks