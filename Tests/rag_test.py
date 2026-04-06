import os
import requests
import chromadb
from sentence_transformers import SentenceTransformer

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
BASE_PATH        = "C:\\RAG Test"               # Folder to scan for .txt files
DB_PATH          = "C:\\RAG Test\\rag_db"        # Persistent ChromaDB storage location
COLLECTION_NAME  = "knowledge"
CHUNK_SIZE       = 400               # Characters per chunk
CHUNK_OVERLAP    = 80                # Overlap between chunks
TOP_K            = 3                 # Number of chunks to retrieve
MIN_SCORE        = 1.2               # Relevance threshold
OLLAMA_URL       = "http://localhost:11434/api/generate"
OLLAMA_MODEL     = "mistral"         # The model we downloaded

# ─────────────────────────────────────────────
# INIT
# ─────────────────────────────────────────────
print("🔧 Loading embedding model...")
model = SentenceTransformer("all-MiniLM-L6-v2")

client = chromadb.PersistentClient(path=DB_PATH)
collection = client.get_or_create_collection(name=COLLECTION_NAME)

# ─────────────────────────────────────────────
# CHUNKING
# ─────────────────────────────────────────────
def chunk_text(text, chunk_size=CHUNK_SIZE, overlap=CHUNK_OVERLAP):
    """Split text into overlapping chunks for better retrieval."""
    chunks = []
    start = 0
    while start < len(text):
        end = start + chunk_size
        chunks.append(text[start:end].strip())
        start += chunk_size - overlap
    return [c for c in chunks if len(c) > 50]

# ─────────────────────────────────────────────
# LOAD & INDEX DOCUMENTS
# ─────────────────────────────────────────────
def load_documents():
    existing_ids = set(collection.get()["ids"])
    new_chunks, new_embeddings, new_ids, new_metadata = [], [], [], []

    for root, dirs, files in os.walk(BASE_PATH):
        dirs[:] = [d for d in dirs if d != "rag_db"]
        for file in files:
            if not file.endswith(".txt"):
                continue

            filepath = os.path.join(root, file)
            with open(filepath, "r", encoding="utf-8") as f:
                content = f.read()

            chunks = chunk_text(content)
            for i, chunk in enumerate(chunks):
                chunk_id = f"{filepath}::chunk{i}"
                if chunk_id in existing_ids:
                    continue

                new_chunks.append(chunk)
                new_ids.append(chunk_id)
                new_metadata.append({
                    "source": filepath,
                    "chunk_index": i,
                    "filename": file
                })

    if new_chunks:
        print(f"⚙️  Embedding {len(new_chunks)} new chunk(s)...")
        new_embeddings = model.encode(new_chunks).tolist()
        collection.add(
            documents=new_chunks,
            embeddings=new_embeddings,
            ids=new_ids,
            metadatas=new_metadata
        )
        print(f"✅ {len(new_chunks)} chunk(s) added to knowledge base.")
    else:
        print("✅ Knowledge base is up to date. No new documents found.")

# ─────────────────────────────────────────────
# RETRIEVE RELEVANT CHUNKS
# ─────────────────────────────────────────────
def retrieve(question):
    """Find the most relevant chunks from the knowledge base."""
    query_embedding = model.encode([question]).tolist()
    results = collection.query(
        query_embeddings=query_embedding,
        n_results=TOP_K,
        include=["documents", "metadatas", "distances"]
    )

    docs      = results["documents"][0]
    metas     = results["metadatas"][0]
    distances = results["distances"][0]

    filtered = []
    for doc, meta, dist in zip(docs, metas, distances):
        if dist <= MIN_SCORE:
            filtered.append({
                "text": doc.strip(),
                "source": meta.get("source", "Unknown"),
                "filename": meta.get("filename", "Unknown")
            })

    return filtered

# ─────────────────────────────────────────────
# ASK MISTRAL (via Ollama)
# ─────────────────────────────────────────────
def ask_mistral(question, chunks):
    """Send the question + retrieved context to Mistral for a natural answer."""

    if not chunks:
        return (
            "I couldn't find anything relevant in my knowledge base.\n"
            "Try rephrasing your question or check that related documents are loaded."
        ), []

    # Build the context block from retrieved chunks
    context = "\n\n".join([f"[Source: {c['filename']}]\n{c['text']}" for c in chunks])
    sources = list(set([c["source"] for c in chunks]))

    # This is the "prompt" — instructions we give Mistral before it answers
    prompt = f"""You are Libby, an offline survival knowledge assistant.
Your job is to answer the user's question using ONLY the information provided below.
Do not use any outside knowledge. Do not guess or make things up.
If the provided information does not contain a clear answer, say so honestly.
Keep your answer clear, practical, and helpful.

--- KNOWLEDGE BASE CONTEXT ---
{context}
--- END OF CONTEXT ---

User question: {question}

Answer:"""

    try:
        # Send the prompt to Ollama running locally
        # "localhost" means "this computer" — no internet involved
        response = requests.post(
            OLLAMA_URL,
            json={
                "model": OLLAMA_MODEL,
                "prompt": prompt,
                "stream": False       # Wait for full response before returning
            },
            timeout=60                # Wait up to 60 seconds for a response
        )
        response.raise_for_status()
        answer = response.json().get("response", "").strip()
        return answer, sources

    except requests.exceptions.ConnectionError:
        return (
            "⚠️  Ollama is not running.\n"
            "    Open a NEW CMD window and type:  ollama serve\n"
            "    Then try your question again."
        ), []

    except Exception as e:
        return f"⚠️  Unexpected error talking to Mistral: {str(e)}", []

# ─────────────────────────────────────────────
# FORMAT OUTPUT
# ─────────────────────────────────────────────
def format_answer(answer, sources):
    lines = []
    lines.append("\n" + "═" * 55)
    lines.append("  LIBBY'S ANSWER")
    lines.append("═" * 55)

    # Word-wrap the answer at 75 characters
    words = answer.split()
    line = []
    for word in words:
        line.append(word)
        if len(" ".join(line)) > 75:
            lines.append("  " + " ".join(line[:-1]))
            line = [word]
    if line:
        lines.append("  " + " ".join(line))

    if sources:
        lines.append("\n" + "─" * 55)
        lines.append("  📄 Sources consulted:")
        for s in sources:
            lines.append(f"     • {s}")

    lines.append("═" * 55)
    return "\n".join(lines)

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
if __name__ == "__main__":
    print("\n🧠 Libby — Offline Knowledge Assistant")
    print("=" * 55)

    load_documents()

    total = collection.count()
    print(f"📦 Total chunks in knowledge base: {total}")
    print("\nType your question below. Type 'exit' to quit.\n")

    while True:
        question = input("❓ Question: ").strip()
        if not question:
            continue
        if question.lower() == "exit":
            print("👋 Goodbye. Knowledge base saved.")
            break

        print("\n⏳ Libby is thinking...\n")
        chunks = retrieve(question)
        answer, sources = ask_mistral(question, chunks)
        print(format_answer(answer, sources))
        print()