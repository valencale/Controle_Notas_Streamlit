import streamlit as st
import pandas as pd
import numpy as np
import os

# ══════════════════════════════════════════════════════════════
# LAZY IMPORTS — sentence_transformers e faiss são pesados.
# Só carregamos quando realmente necessário para não travar
# o app caso o pacote não esteja instalado no ambiente ativo.
# ══════════════════════════════════════════════════════════════

def _get_encoder():
    """Carrega (com cache) o modelo de embeddings."""
    from sentence_transformers import SentenceTransformer
    return SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')


@st.cache_resource(show_spinner=False)
def _load_encoder():
    return _get_encoder()


def _load_or_create_faiss_index(base_dir):
    import faiss

    texts_path = os.path.join(base_dir, "data", "rag_texts.npy")
    index_path = os.path.join(base_dir, "data", "rag_index.faiss")

    if not os.path.exists(index_path) or not os.path.exists(texts_path):
        encoder = _load_encoder()
        dimension = encoder.get_sentence_embedding_dimension()
        index = faiss.IndexFlatL2(dimension)
        return index, []

    index = faiss.read_index(index_path)
    texts = np.load(texts_path, allow_pickle=True).tolist()
    return index, texts


# ══════════════════════════════════════════════════════════════
# INGESTÃO DE DADOS
# ══════════════════════════════════════════════════════════════

def _ingest_dataframe(df, source_name):
    """Converte cada linha de um DataFrame em um texto legível para RAG."""
    chunks = []
    df = df.fillna("N/A")
    for _, row in df.iterrows():
        row_text = f"Fonte: {source_name}. " + " | ".join(
            [f"{col}: {val}" for col, val in row.items()]
        )
        chunks.append(row_text)
    return chunks


def _ingest_pdf(pdf_path):
    """Extrai texto de um PDF usando pdfplumber."""
    import pdfplumber

    text_chunks = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text:
                    text_chunks.append(
                        f"Fonte: PDF {os.path.basename(pdf_path)}, "
                        f"Página {page_num + 1}. Conteúdo: {text}"
                    )
    except Exception as e:
        print(f"Erro ao ler PDF {pdf_path}: {e}")
    return text_chunks


# ══════════════════════════════════════════════════════════════
# CONSTRUÇÃO DO ÍNDICE
# ══════════════════════════════════════════════════════════════

def rebuild_index(base_dir):
    """Reconstrói todo o índice FAISS a partir das planilhas e PDFs."""
    import faiss

    chunks = []

    # 1. Planilhas Excel
    excel_files = [
        "CONTROLE NOTAS.xlsm",
        "RELATÓRIO DE ENTREGAS 2026_LEVE.xlsb",
        "expedicoes.xlsx",
    ]
    for file in excel_files:
        path = os.path.join(base_dir, file)
        if os.path.exists(path):
            try:
                engine = "pyxlsb" if path.endswith(".xlsb") else "openpyxl"
                df = pd.read_excel(path, engine=engine)
                chunks.extend(_ingest_dataframe(df, file))
            except Exception as e:
                print(f"Erro ao ler {file}: {e}")

    # 2. PDFs da raiz do projeto
    for file in os.listdir(base_dir):
        if file.lower().endswith(".pdf"):
            chunks.extend(_ingest_pdf(os.path.join(base_dir, file)))

    if not chunks:
        return False

    encoder = _load_encoder()
    embeddings = encoder.encode(chunks, convert_to_numpy=True)

    dimension = embeddings.shape[1]
    index = faiss.IndexFlatL2(dimension)
    index.add(embeddings)

    data_dir = os.path.join(base_dir, "data")
    os.makedirs(data_dir, exist_ok=True)
    faiss.write_index(index, os.path.join(data_dir, "rag_index.faiss"))
    np.save(os.path.join(data_dir, "rag_texts.npy"), np.array(chunks, dtype=object))

    return True


# ══════════════════════════════════════════════════════════════
# BUSCA SEMÂNTICA
# ══════════════════════════════════════════════════════════════

def search_rag(query, k=5, base_dir=None):
    """Busca os k trechos mais relevantes no índice FAISS."""
    if base_dir is None:
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    index, texts = _load_or_create_faiss_index(base_dir)

    if index.ntotal == 0 or not texts:
        return [
            "A base de conhecimento está vazia. "
            "O administrador precisa reconstruir o índice (Sidebar → Administração)."
        ]

    encoder = _load_encoder()
    query_vector = encoder.encode([query], convert_to_numpy=True)
    distances, indices = index.search(query_vector, k)

    results = []
    for idx in indices[0]:
        if 0 <= idx < len(texts):
            results.append(texts[idx])

    return results if results else ["Nenhum resultado relevante encontrado."]
