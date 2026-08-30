"""Chunking, SQLite persistence, and BM25 retrieval per indexed folder."""

import hashlib
import re
import sqlite3
import threading
import time
from pathlib import Path

from rank_bm25 import BM25Okapi

from .box_bridge import CACHE_BOX_DIR, PROJECT_DIR
from .extract import extract

INDEX_DIR = PROJECT_DIR / ".cache" / "index"

CHUNK_SIZE = 1500
CHUNK_OVERLAP = 200

_bm25_cache: dict[str, tuple[float, BM25Okapi, list[dict]]] = {}
_cache_lock = threading.Lock()


def make_slug(box_path: str) -> str:
    name = box_path.rstrip("\\").split("\\")[-1]
    safe = re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")[:40] or "folder"
    digest = hashlib.sha256(box_path.encode("utf-8")).hexdigest()[:8]
    return f"{safe}-{digest}"


def db_path(slug: str) -> Path:
    return INDEX_DIR / f"{slug}.db"


def _connect(slug: str) -> sqlite3.Connection:
    INDEX_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path(slug))
    conn.row_factory = sqlite3.Row
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS files(
            rel_path TEXT PRIMARY KEY, size INTEGER, mtime REAL, ext TEXT,
            status TEXT, n_chunks INTEGER, note TEXT);
        CREATE TABLE IF NOT EXISTS chunks(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rel_path TEXT, loc TEXT, text TEXT);
        CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY, value TEXT);
        CREATE INDEX IF NOT EXISTS chunks_rel ON chunks(rel_path);
    """)
    return conn


def tokenize(text: str) -> list[str]:
    return re.findall(r"[a-z0-9]{2,}", text.lower())


def chunk_text(text: str, size: int = CHUNK_SIZE, overlap: int = CHUNK_OVERLAP) -> list[str]:
    """Split on sentence/newline boundaries into ~size-char chunks with overlap."""
    if len(text) <= size:
        return [text]
    pieces = re.split(r"(?<=[.!?])\s+|\n+", text)
    chunks, buf = [], ""
    for piece in pieces:
        if not piece:
            continue
        if len(buf) + len(piece) + 1 > size and buf:
            chunks.append(buf)
            buf = buf[-overlap:] + " " if overlap else ""
        buf += piece + " "
    if buf.strip():
        chunks.append(buf)
    return [c.strip() for c in chunks if c.strip()]


def index_folder(slug: str, box_path: str, progress_cb=None) -> dict:
    """Extract + chunk every synced file whose (size, mtime) changed; prune vanished files."""
    cache_dir = CACHE_BOX_DIR / slug
    conn = _connect(slug)
    try:
        manifest = {r["rel_path"]: r for r in conn.execute("SELECT * FROM files")}
        on_disk = [p for p in cache_dir.rglob("*") if p.is_file()]

        done, changed = 0, 0
        for path in on_disk:
            rel = str(path.relative_to(cache_dir))
            stat = path.stat()
            done += 1
            row = manifest.pop(rel, None)
            if row is not None and row["size"] == stat.st_size and abs(row["mtime"] - stat.st_mtime) < 2:
                continue  # unchanged
            changed += 1
            if progress_cb:
                progress_cb({"done": done, "total": len(on_disk), "current_file": rel})

            result = extract(path)
            conn.execute("DELETE FROM chunks WHERE rel_path = ?", (rel,))
            n_chunks = 0
            for section in result.sections:
                for chunk in chunk_text(section.text):
                    conn.execute(
                        "INSERT INTO chunks(rel_path, loc, text) VALUES (?,?,?)",
                        (rel, section.loc, chunk))
                    n_chunks += 1
            conn.execute(
                "INSERT OR REPLACE INTO files(rel_path, size, mtime, ext, status, n_chunks, note) "
                "VALUES (?,?,?,?,?,?,?)",
                (rel, stat.st_size, stat.st_mtime, path.suffix.lower(),
                 result.status, n_chunks, result.note))
            conn.commit()

        # Files present in the manifest but gone from disk
        for rel in manifest:
            conn.execute("DELETE FROM chunks WHERE rel_path = ?", (rel,))
            conn.execute("DELETE FROM files WHERE rel_path = ?", (rel,))

        conn.execute("INSERT OR REPLACE INTO meta VALUES ('box_path', ?)", (box_path,))
        conn.execute("INSERT OR REPLACE INTO meta VALUES ('indexed_at', ?)", (str(time.time()),))
        conn.commit()

        counts = dict(conn.execute("SELECT status, COUNT(*) FROM files GROUP BY status"))
        n_chunks = conn.execute("SELECT COUNT(*) FROM chunks").fetchone()[0]
        return {"files": len(on_disk), "changed": changed, "chunks": n_chunks, "by_status": counts}
    finally:
        conn.close()


def load_index(slug: str) -> tuple[BM25Okapi, list[dict]] | None:
    """BM25 over all chunks, cached in-process and invalidated by DB mtime."""
    path = db_path(slug)
    if not path.exists():
        return None
    mtime = path.stat().st_mtime
    with _cache_lock:
        cached = _bm25_cache.get(slug)
        if cached and cached[0] == mtime:
            return cached[1], cached[2]

    conn = _connect(slug)
    try:
        refs = [
            {"rel_path": r["rel_path"], "loc": r["loc"], "text": r["text"], "status": r["status"]}
            for r in conn.execute(
                "SELECT c.rel_path, c.loc, c.text, f.status FROM chunks c "
                "JOIN files f ON f.rel_path = c.rel_path")
        ]
    finally:
        conn.close()
    if not refs:
        return None
    bm25 = BM25Okapi([tokenize(r["text"]) for r in refs])
    with _cache_lock:
        _bm25_cache[slug] = (mtime, bm25, refs)
    return bm25, refs


def folder_info(slug: str) -> dict | None:
    path = db_path(slug)
    if not path.exists():
        return None
    conn = _connect(slug)
    try:
        meta = dict(conn.execute("SELECT key, value FROM meta"))
        by_status = dict(conn.execute("SELECT status, COUNT(*) FROM files GROUP BY status"))
        n_files = conn.execute("SELECT COUNT(*) FROM files").fetchone()[0]
        n_chunks = conn.execute("SELECT COUNT(*) FROM chunks").fetchone()[0]
        return {
            "slug": slug, "box_path": meta.get("box_path", ""),
            "name": meta.get("box_path", slug).rstrip("\\").split("\\")[-1],
            "indexed_at": float(meta["indexed_at"]) if "indexed_at" in meta else None,
            "files": n_files, "chunks": n_chunks, "by_status": by_status,
        }
    finally:
        conn.close()


def list_folders() -> list[dict]:
    if not INDEX_DIR.exists():
        return []
    infos = [folder_info(p.stem) for p in sorted(INDEX_DIR.glob("*.db"))]
    return [i for i in infos if i]


def folder_files(slug: str) -> list[dict]:
    conn = _connect(slug)
    try:
        return [dict(r) for r in conn.execute("SELECT * FROM files ORDER BY rel_path")]
    finally:
        conn.close()
