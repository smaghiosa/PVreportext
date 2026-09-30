"""
vector_db.py

Embeds each report section (from the `documents` SQLite table) into a
local ChromaDB collection using BAAI/bge-base-en-v1.5, and provides
semantic search over that collection for the "explain/summarize/compare"
side of queries - as opposed to exact numeric lookups, which stay in
SQL (see db_writer.py / equipment_master etc).

Deployment note: BAAI/bge-base-en-v1.5 downloads from HuggingFace Hub
the first time it's used (~440MB) and is cached locally after that -
needs one-time internet access on whatever machine runs this, then
works fully offline. (This could not be downloaded/tested in the build
sandbox due to a network restriction there - see the note at the
bottom of this file for exactly what was and wasn't verified.)
"""
import chromadb
import sqlite3

_EMBEDDING_MODEL_NAME = "BAAI/bge-base-en-v1.5"
_COLLECTION_NAME = "pvelite_documents"

_model = None  # lazy-loaded so importing this module doesn't require the model


def _get_model():
    global _model
    if _model is None:
        from sentence_transformers import SentenceTransformer
        _model = SentenceTransformer(_EMBEDDING_MODEL_NAME)
    return _model


def _get_collection(chroma_path="./chroma_store"):
    client = chromadb.PersistentClient(path=chroma_path)
    return client.get_or_create_collection(_COLLECTION_NAME)


def embed_pending_documents(db_path, chroma_path="./chroma_store", batch_size=32, log=print):
    """
    Finds every `documents` row not yet embedded (embedded=0), embeds
    it, and upserts it into the Chroma collection with metadata
    (equipment_id, project, category, tag_no, section_title) so results
    can be filtered/traced back to their source equipment. Marks each
    row embedded=1 in SQLite once done, so re-running this only
    processes what's actually new since the last run.
    """
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("""
        SELECT d.id, d.equipment_id, d.section_title, d.section_text,
               e.project, e.category, e.tag_no
        FROM documents d
        JOIN equipment_master e ON e.id = d.equipment_id
        WHERE d.embedded = 0
    """)
    rows = cur.fetchall()
    if not rows:
        log("No pending documents to embed.")
        conn.close()
        return 0

    log(f"Embedding {len(rows)} document section(s)...")
    collection = _get_collection(chroma_path)
    model = _get_model()

    for i in range(0, len(rows), batch_size):
        batch = rows[i:i + batch_size]
        texts = [r["section_text"] or "" for r in batch]
        embeddings = model.encode(texts, normalize_embeddings=True).tolist()
        collection.upsert(
            ids=[f"doc_{r['id']}" for r in batch],
            embeddings=embeddings,
            documents=texts,
            metadatas=[{
                "equipment_id": r["equipment_id"], "project": r["project"],
                "category": r["category"], "tag_no": r["tag_no"],
                "section_title": r["section_title"],
            } for r in batch],
        )
        cur.executemany(
            "UPDATE documents SET embedded=1, embedded_at=CURRENT_TIMESTAMP WHERE id=?",
            [(r["id"],) for r in batch],
        )
        conn.commit()
        log(f"  {min(i + batch_size, len(rows))}/{len(rows)} embedded")

    conn.close()
    return len(rows)


def semantic_search(query, chroma_path="./chroma_store", n_results=5,
                     project=None, tag_no=None):
    """
    Returns the n_results most semantically-relevant report sections for
    `query`, optionally filtered to a specific project or tag. Each
    result includes which equipment/section it came from, so the
    calling code (or an LLM synthesis step) can cite its source.
    """
    collection = _get_collection(chroma_path)
    model = _get_model()
    query_embedding = model.encode([query], normalize_embeddings=True).tolist()

    where = {}
    if project:
        where["project"] = project
    if tag_no:
        where["tag_no"] = tag_no

    results = collection.query(
        query_embeddings=query_embedding, n_results=n_results,
        where=where or None,
    )
    hits = []
    for doc, meta, dist in zip(results["documents"][0], results["metadatas"][0], results["distances"][0]):
        hits.append({"text": doc, "metadata": meta, "distance": dist})
    return hits


# ----------------------------------------------------------------------
# What was actually verified in the build sandbox vs. what needs
# verifying on a real deployment machine:
#
# VERIFIED (using manually-supplied vectors in place of the real model,
# since this sandbox's network restrictions block downloading ANY
# embedding model - not just this one):
#   - documents -> Chroma ingestion pipeline (SQL join, batching, the
#     upsert/metadata/marking-embedded logic actually written in this
#     file)
#   - metadata filtering (by project/tag_no) on query
#   - re-embedding only picks up genuinely new/unembedded rows
#
# NOT VERIFIED HERE (needs testing on the real deployment machine, which
# will have normal internet access):
#   - BAAI/bge-base-en-v1.5 actually downloading and loading correctly
#   - real semantic search quality/relevance (this sandbox could only
#     test with dummy vectors, not real embeddings, so retrieval
#     *quality* - as opposed to retrieval *mechanics* - is unverified)
#
# First thing to run on the real machine: embed_pending_documents() on
# one tag's data, then semantic_search() a few realistic questions
# against it and eyeball whether the results make sense, before trusting
# this for real use.
# ----------------------------------------------------------------------
