"""Critical-dates tracker.

For properties at LOI or later, download just their deal documents (LOI, PSA,
estoppel, settlement…), have `claude -p` read them, and extract the Critical
Dates from the user's Deal Timeline (effective date, inspection/DD end, deposit
goes non-refundable, closing date, estoppel/notice windows). Cached per property
in .cache/status/dates.json.
"""

import json
import os
import re
import signal
import subprocess
import threading

from . import indexer, jobs, status
from .box_bridge import PROJECT_DIR, _b64, _ps_args, wsl_to_win

DATES_PATH = status.STATUS_DIR / "dates.json"
DOCS_CACHE = PROJECT_DIR / ".cache" / "box" / "_dates"

_lock = threading.Lock()

# Which files are worth reading, in priority order. A filled-in Critical Dates
# Timeline (the brokerage's own CDT form) is the best source when it exists.
DOC_PATTERNS = [
    ("cdt", re.compile(r"critical\s+dates|\bcdt\b", re.I)),
    ("psa", re.compile(r"\bpsa\b|purchase\s+(and|&)\s+sale|purchase\s+agreement|\bcontract\b", re.I)),
    ("loi", re.compile(r"\bloi\b|letter\s+of\s+intent", re.I)),
    ("settlement", re.compile(r"settlement|closing\s+statement", re.I)),
    ("estoppel", re.compile(r"estoppel", re.I)),
]
READABLE_EXTS = {".pdf", ".docx", ".txt"}  # .doc (legacy) has no extractor
MAX_DOCS = 5
MAX_BYTES = 15 * 1024 * 1024

DATE_FIELDS = [
    ("effective_date", "Effective date"),
    ("deposit_date", "Deposit due"),
    ("inspection_end", "DD / inspection ends"),
    ("deposit_nonrefundable", "Deposit non-refundable"),
    ("financing_contingency", "Financing contingency"),
    ("closing_date", "Closing date"),
]

PROMPT = """You are extracting critical deal dates from commercial real estate documents.
Read each of these files with the Read tool:
{files}

Return ONLY a JSON object (no prose, no markdown fence) with these keys, using
"YYYY-MM-DD" for dates you find and null for anything not stated. If a date is
defined relative to another (e.g. "30 days after the Effective Date") and you can
compute it, do so; otherwise put the rule text in notes.
{{
  "effective_date": null,
  "deposit_date": null,
  "inspection_end": null,
  "deposit_nonrefundable": null,
  "financing_contingency": null,
  "closing_date": null,
  "estoppel_days": null,
  "notice_days": null,
  "notes": "one short sentence with anything important, or empty string"
}}"""


def _load() -> dict:
    try:
        return json.loads(DATES_PATH.read_text())
    except (OSError, json.JSONDecodeError):
        return {}


def _save(data: dict) -> None:
    status.STATUS_DIR.mkdir(parents=True, exist_ok=True)
    DATES_PATH.write_text(json.dumps(data, indent=1))


def candidate_docs(files: list[dict]) -> list[str]:
    picked, seen = [], set()
    clean = status.deal_files(files)  # same template/asset noise filter as the checker
    for _, pattern in DOC_PATTERNS:
        for f in clean:
            name, base = f["n"], f["base"]
            ext = ("." + base.rsplit(".", 1)[-1].lower()) if "." in base else ""
            if name in seen or ext not in READABLE_EXTS:
                continue
            if pattern.search(base):
                picked.append(name)
                seen.add(name)
                if len(picked) >= MAX_DOCS:
                    return picked
    return picked


def _readable_copy(local_path: str) -> str | None:
    """Claude's Read tool handles PDF/text but not .docx — convert those to a .txt sidecar."""
    from pathlib import Path
    from .extract import extract
    p = Path(local_path)
    if p.suffix.lower() != ".docx":
        return local_path
    result = extract(p)
    text = "\n\n".join(s.text for s in result.sections)
    if not text.strip():
        return None
    txt = p.with_suffix(".docx.txt")
    txt.write_text(text)
    return str(txt)


def _sync_docs(box_path: str, slug: str, rel_paths: list[str]) -> list[str]:
    """Copy the chosen files locally; returns WSL paths of the copies."""
    dest_root = DOCS_CACHE / slug
    dest_root.mkdir(parents=True, exist_ok=True)
    dest_root_win = wsl_to_win(dest_root)
    pairs = []
    for rel in rel_paths:
        src = f"{box_path}\\{rel}"
        pairs.append(_b64(src) + "|" + _b64(dest_root_win + "\\" + rel))
    proc = subprocess.run(
        _ps_args("sync_files.ps1"), input="\n".join(pairs) + "\n",
        capture_output=True, text=True, timeout=600,
    )
    ok_paths = []
    for line in proc.stdout.splitlines():
        line = line.strip().lstrip("﻿")
        if not line:
            continue
        try:
            r = json.loads(line)
        except json.JSONDecodeError:
            continue
        if r.get("ok"):
            rel = r["src"][len(box_path):].lstrip("\\")
            ok_paths.append(str(dest_root / rel.replace("\\", "/")))
    return ok_paths


def _extract_with_claude(local_paths: list[str]) -> dict:
    prompt = PROMPT.format(files="\n".join(f"- {p}" for p in local_paths))
    cmd = [
        "claude", "-p", "--output-format", "json",
        "--model", os.environ.get("QA_MODEL", "sonnet"),
        "--no-session-persistence", "--tools", "Read",
        "--add-dir", str(DOCS_CACHE),
        "--permission-mode", "dontAsk",
    ]
    proc = subprocess.Popen(cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, text=True, start_new_session=True)
    try:
        stdout, stderr = proc.communicate(input=prompt, timeout=240)
    except subprocess.TimeoutExpired:
        os.killpg(os.getpgid(proc.pid), signal.SIGKILL)
        proc.wait()
        raise RuntimeError("claude timed out")
    if proc.returncode != 0:
        raise RuntimeError(f"claude exited {proc.returncode}: {stderr.strip()[:300]}")
    envelope = json.loads(stdout)
    if envelope.get("is_error"):
        raise RuntimeError(str(envelope.get("result"))[:300])
    text = envelope["result"]
    match = re.search(r"\{.*\}", text, re.S)
    if not match:
        raise RuntimeError(f"no JSON in claude output: {text[:200]}")
    return json.loads(match.group(0))


def extract_property(box_path: str, entry: dict) -> dict:
    files = entry.get("files") or []
    docs = candidate_docs(files)
    if not docs:
        return {"skipped": "no readable deal documents found"}
    slug = indexer.make_slug(box_path)
    local = _sync_docs(box_path, slug, docs)
    readable = [r for r in (_readable_copy(p) for p in local) if r]
    if not readable:
        return {"skipped": "documents could not be copied or read"}
    dates = _extract_with_claude(readable)
    dates["source_files"] = docs
    return dates


def run_extract(job_id: str, force: bool = False) -> dict:
    import time
    scan = status._load(status.SCAN_PATH, {})
    overrides = status._load(status.OVERRIDES_PATH, {})
    targets = []
    for path, entry in scan.items():
        if "files" not in entry:
            continue
        info = status.classify(entry["files"])
        idx = info["auto_stage_idx"]
        if overrides.get(path) in status.STAGE_KEYS:
            idx = status.STAGE_KEYS.index(overrides[path]) + 1
        if idx >= 3:  # LOI or later
            targets.append((path, entry))

    jobs.update(job_id, phase="extract-dates", total=len(targets), done=0)
    cache = _load()
    # Drop results for properties that no longer qualify (stage rules changed, folder gone)
    target_paths = {p for p, _ in targets}
    cache = {p: v for p, v in cache.items() if p in target_paths}
    done, errors = 0, 0
    for path, entry in targets:
        done += 1
        jobs.update(job_id, done=done, current_file=entry.get("name", ""))
        existing = cache.get(path)
        if (existing and not force and not existing.get("error")
                and existing.get("scan_time") == entry.get("scanned_at")):
            continue  # unchanged since last successful extraction
        try:
            result = extract_property(path, entry)
        except Exception as e:  # noqa: BLE001 - keep going per property
            errors += 1
            result = {"error": f"{type(e).__name__}: {e}"[:300]}
        result["name"] = entry.get("name")
        result["extracted_at"] = time.time()
        result["scan_time"] = entry.get("scanned_at")
        if existing:  # keep dates that were approved from email over re-extraction
            for field in existing.get("locked", []):
                result[field] = existing.get(field)
            for k in ("locked", "sources"):
                if existing.get(k):
                    result[k] = existing[k]
        cache[path] = result
        with _lock:
            _save(cache)
    with _lock:
        _save(cache)  # persist pruning even when every target was skipped as unchanged
    return {"extracted": done, "errors": errors}


def start_extract(force: bool = False) -> str:
    job_id = jobs.create("extract-dates")

    def work():
        result = run_extract(job_id, force=force)
        jobs.update(job_id, message=f"Extracted dates for {result['extracted']} "
                                    f"properties ({result['errors']} errors).")

    jobs.run_in_thread(job_id, work)
    return job_id


def get_dates() -> dict:
    return {"fields": [{"key": k, "label": l} for k, l in DATE_FIELDS],
            "properties": _load()}
