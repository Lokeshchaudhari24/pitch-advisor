import streamlit as st
import chromadb


@st.cache_resource
def get_collection():
    client = chromadb.PersistentClient(
        path="chroma_db"
    )

    return client.get_or_create_collection(
        name="policy_documents"
    )


collection = get_collection()


def store_chunks(chunks, embeddings):

    ids = [
        f"chunk_{i}"
        for i in range(len(chunks))
    ]

    texts = [
        chunk["text"]
        for chunk in chunks
    ]

    metadatas = [
        chunk["metadata"]
        for chunk in chunks
    ]

    collection.add(
        ids=ids,
        documents=texts,
        embeddings=embeddings.tolist(),
        metadatas=metadatas
    )

    print(f"Stored {len(chunks)} chunks.")