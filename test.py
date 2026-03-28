import chromadb
import os

BASE_PATH = "C:\\RAG Test"
DB_PATH = "C:\\RAG Test\\rag_db"

client = chromadb.PersistentClient(path=DB_PATH)
collection = client.get_or_create_collection(name="knowledge")

print("Files found:")
for root, dirs, files in os.walk(BASE_PATH):
    dirs[:] = [d for d in dirs if "rag_db" not in d]
    for f in files:
        if f.endswith(".txt") or f.endswith(".xlsx"):
            print(" -", os.path.join(root, f))

print("Chunks in DB:", collection.count())
