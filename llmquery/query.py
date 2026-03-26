#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.12"
# dependencies = ["chromadb", "ollama"]
# ///
import argparse, chromadb, ollama

MODEL_EMBED = "nomic-embed-text"
MODEL_CHAT  = "llama3.2"

parser = argparse.ArgumentParser(description="Query PST emails via RAG")
parser.add_argument("query", nargs="+", help="Question to ask")
parser.add_argument("--collection", default="testPST", help="ChromaDB collection name (default: testPST)")
parser.add_argument("--n-results", type=int, default=5, metavar="N", help="Number of emails to retrieve (default: 5)")
args = parser.parse_args()

question = " ".join(args.query)

client = chromadb.HttpClient(host="localhost", port=8000)
collection = client.get_collection(args.collection)

q_embedding = ollama.embeddings(model=MODEL_EMBED, prompt=question)["embedding"]

results = collection.query(query_embeddings=[q_embedding], n_results=args.n_results)
docs = results["documents"][0]
metas = results["metadatas"][0]

context = "\n\n---\n\n".join(
    f"From: {m['from']}\nDate: {m['date']}\nSubject: {m['subject']}\n\n{d}"
    for d, m in zip(docs, metas)
)

response = ollama.chat(model=MODEL_CHAT, messages=[{
    "role": "user",
    "content": f"Answer using only these emails:\n\n{context}\n\nQuestion: {question}"
}])
print(response["message"]["content"])
