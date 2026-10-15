import json
import os
import numpy as np

KNOWLEDGE_PATH = os.getenv("KNOWLEDGE_PATH", "knowledge.json")

with open(KNOWLEDGE_PATH, encoding="utf-8") as f:
    KNOWLEDGE = json.load(f)


DEFAULT_MIN_SCORE = float(os.getenv("RAG_MIN_SCORE", "0.5"))


def cosine(a, b):

    try:
        a = np.array(a, dtype=float)
        b = np.array(b, dtype=float)

        if a.shape != b.shape:
            print(f"[retrieval] WARNING: embedding shape mismatch "
                  f"query={a.shape} doc={b.shape} - check that ingestion "
                  f"and query embeddings use the same model/config.")
            return 0.0

        norm_a = np.linalg.norm(a)
        norm_b = np.linalg.norm(b)

        if norm_a == 0 or norm_b == 0:
            return 0.0

        return float(np.dot(a, b) / (norm_a * norm_b))
    except Exception as e:
        print(f"[retrieval] cosine() failed: {e}")
        return 0.0


def search(query_embedding, top_k=5, min_score=None, max_chunks_per_source=2):

    min_score = DEFAULT_MIN_SCORE if min_score is None else min_score

    raw_results = []
    for item in KNOWLEDGE:
        score = cosine(query_embedding, item["embedding"])
        raw_results.append({
            "score": score,
            "source": item["source"],
            "text": item["text"],
        })

    raw_results.sort(key=lambda x: x["score"], reverse=True)

    filtered_results = []
    source_counts = {}

    for item in raw_results:
        if item["score"] < min_score:
            break

        src = item["source"]
        count = source_counts.get(src, 0)

        if count < max_chunks_per_source:
            filtered_results.append(item)
            source_counts[src] = count + 1

        if len(filtered_results) >= top_k:
            break

    if not filtered_results and raw_results and raw_results[0]["score"] > 0:
        fallback = dict(raw_results[0])
        fallback["low_confidence"] = True
        filtered_results = [fallback]

    return filtered_results


def keyword_search(needle: str, max_results: int = 8):

    if not needle:
        return []
    needle_upper = needle.upper()
    hits = []
    for item in KNOWLEDGE:
        if needle_upper in item["text"].upper():
            hits.append({
                "score": 1.0,
                "source": item["source"],
                "text": item["text"],
            })
    return hits[:max_results]