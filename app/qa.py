"""Retrieval + prompt build + `claude -p` subprocess for answering questions."""

import json
import os
import signal
import subprocess
from pathlib import Path

from . import indexer
from .box_bridge import CACHE_BOX_DIR

TOP_K = 8
EXCERPT_BUDGET = 12_000
MAX_ATTACH = 2
HISTORY_TURNS = 6
HISTORY_TRUNC = 500
CLAUDE_TIMEOUT = 180
CLAUDE_MODEL = os.environ.get("QA_MODEL", "sonnet")

# per-slug chat history: [(question, answer), ...] — in-memory, single-user app
_history: dict[str, list[tuple[str, str]]] = {}


def retrieve(slug: str, question: str) -> tuple[list[dict], list[str]]:
    """Top text chunks within budget, plus up to MAX_ATTACH attachable file paths."""
    loaded = indexer.load_index(slug)
    if loaded is None:
        return [], []
    bm25, refs = loaded
    scores = bm25.get_scores(indexer.tokenize(question))
    ranked = sorted(zip(scores, refs), key=lambda t: -t[0])

    chunks, attach, used, seen_attach = [], [], 0, set()
    for score, ref in ranked:
        if score <= 0:
            break
        if ref["status"] == "attachable":
            if ref["rel_path"] not in seen_attach and len(attach) < MAX_ATTACH:
                seen_attach.add(ref["rel_path"])
                attach.append(ref["rel_path"])
            continue
        if len(chunks) >= TOP_K or used + len(ref["text"]) > EXCERPT_BUDGET:
            continue
        chunks.append({**ref, "score": float(score)})
        used += len(ref["text"])
    return chunks, attach


def build_prompt(slug: str, question: str, chunks: list[dict], attach: list[str]) -> str:
    cache_dir = CACHE_BOX_DIR / slug
    parts = [
        "You are a document Q&A assistant for a folder of business documents.",
        "Answer the question using ONLY the numbered excerpts below (and the listed",
        "attachable files, if any). Cite every claim inline as [filename, loc] using",
        "the FILE and LOC values of the excerpts you relied on. If the excerpts do not",
        "contain the answer, say so plainly - do not guess.",
    ]
    if attach:
        parts.append(
            "\nThese files are images/scans with no extracted text. If they look relevant "
            "to the question, open them with the Read tool and use what you see:")
        parts.extend(f"- {cache_dir / rel}" for rel in attach)

    history = _history.get(slug, [])[-HISTORY_TURNS:]
    if history:
        parts.append("\nRecent conversation (for context only):")
        for q, a in history:
            parts.append(f"Q: {q[:HISTORY_TRUNC]}")
            parts.append(f"A: {a[:HISTORY_TRUNC]}")

    parts.append("\nEXCERPTS:")
    for i, c in enumerate(chunks, 1):
        parts.append(f"[{i}] FILE: {c['rel_path']} | LOC: {c['loc']}\n{c['text']}")

    parts.append(f"\nQUESTION: {question}")
    return "\n".join(parts)


def ask_claude(prompt: str, slug: str) -> str:
    cache_dir = CACHE_BOX_DIR / slug
    cmd = [
        "claude", "-p",
        "--output-format", "json",
        "--model", CLAUDE_MODEL,
        "--no-session-persistence",
        "--tools", "Read",
        "--add-dir", str(cache_dir),
        "--permission-mode", "dontAsk",
    ]
    proc = subprocess.Popen(
        cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, start_new_session=True,
    )
    try:
        stdout, stderr = proc.communicate(input=prompt, timeout=CLAUDE_TIMEOUT)
    except subprocess.TimeoutExpired:
        os.killpg(os.getpgid(proc.pid), signal.SIGKILL)
        proc.wait()
        raise RuntimeError(f"claude timed out after {CLAUDE_TIMEOUT}s")
    if proc.returncode != 0:
        raise RuntimeError(f"claude exited {proc.returncode}: {stderr.strip()[:500]}")
    envelope = json.loads(stdout)
    if envelope.get("is_error"):
        raise RuntimeError(f"claude error: {str(envelope.get('result'))[:500]}")
    return envelope["result"]


def answer(slug: str, question: str) -> dict:
    import time
    start = time.monotonic()
    chunks, attach = retrieve(slug, question)
    if not chunks and not attach:
        return {
            "answer": "This folder's index has no content matching the question "
                      "(it may contain only images or unsupported files).",
            "sources": [], "elapsed": 0.0,
        }
    prompt = build_prompt(slug, question, chunks, attach)
    text = ask_claude(prompt, slug)
    _history.setdefault(slug, []).append((question, text))
    sources = [{"rel_path": c["rel_path"], "loc": c["loc"], "score": round(c["score"], 2)}
               for c in chunks]
    sources += [{"rel_path": rel, "loc": "file", "score": None} for rel in attach]
    return {"answer": text, "sources": sources, "elapsed": round(time.monotonic() - start, 1)}


def clear_history(slug: str) -> None:
    _history.pop(slug, None)


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        sys.exit("usage: python -m app.qa <slug> <question>")
    result = answer(sys.argv[1], sys.argv[2])
    print(result["answer"])
    print("\n--- sources ---")
    for s in result["sources"]:
        print(f"  {s['rel_path']} ({s['loc']}) score={s['score']}")
    print(f"[{result['elapsed']}s]")
