import streamlit as st
from sentence_transformers import SentenceTransformer


@st.cache_resource
def get_embedding_model():
    return SentenceTransformer("BAAI/bge-m3")


def create_embeddings(chunks):
    model = get_embedding_model()

    texts = [chunk["text"] for chunk in chunks]

    embeddings = model.encode(
        texts,
        normalize_embeddings=True
    )

    return embeddings