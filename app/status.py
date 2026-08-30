"""Deal-pipeline status checker.

Stages and checklist items come from the user's "Deal Timeline 1.docx"
(.Finalize Checklist folder). Each property folder is scanned (file names only —
no downloads) and matched against per-item filename evidence. The furthest stage
with evidence is the property's auto status; items expected by that stage but
absent are its "missing docs". Manual stage overrides win over auto detection.
"""

import json
import re
import subprocess
import threading
import time

from . import jobs
from .box_bridge import PROJECT_DIR, _b64, _ps_args, get_box_root, list_children
from .categories import CATEGORY_FOLDERS, category_of, is_property

STATUS_DIR = PROJECT_DIR / ".cache" / "status"
SCAN_PATH = STATUS_DIR / "scan.json"
OVERRIDES_PATH = STATUS_DIR / "overrides.json"

_lock = threading.Lock()

STAGES = [
    {"key": "listing", "label": "Listing"},
    {"key": "marketing", "label": "Prep & Marketing"},
    {"key": "loi", "label": "LOI"},
    {"key": "psa", "label": "PSA / Contract"},
    {"key": "dd", "label": "Due Diligence / Escrow"},
    {"key": "closing", "label": "Closing"},
]
STAGE_KEYS = [s["key"] for s in STAGES]

# Checklist items (from Deal Timeline 1.docx), each tied to a stage.
# A stage counts as reached when ANY of its items has evidence.
ITEMS = [
    {"key": "listing_agreement", "label": "Listing Agreement", "stage": "listing",
     "patterns": [r"listing\s+agreement", r"\blisting\b"]},
    {"key": "lease", "label": "Lease & Amendments", "stage": "marketing",
     "patterns": [r"\blease\b", r"amendment"]},
    {"key": "om", "label": "OM / Marketing Package", "stage": "marketing",
     "patterns": [r"\bom\b", r"offering\s+memo", r"marketing", r"brochure", r"blast", r"flyer"]},
    {"key": "lease_abstract", "label": "Lease Abstract", "stage": "marketing",
     "patterns": [r"lease\s+abstract", r"abstract"]},
    {"key": "financials", "label": "Financials / Rent Roll", "stage": "marketing",
     "patterns": [r"rent\s*roll", r"financial", r"sale\s+comp"]},
    {"key": "loi", "label": "LOI", "stage": "loi",
     "patterns": [r"\bloi\b", r"letter\s+of\s+intent"]},
    {"key": "psa", "label": "PSA / Contract", "stage": "psa",
     "patterns": [r"\bpsa\b", r"purchase\s+(and|&)\s+sale", r"purchase\s+agreement", r"\bcontract\b"]},
    {"key": "estoppel", "label": "Estoppel", "stage": "dd", "patterns": [r"estoppel"]},
    {"key": "snda", "label": "SNDA", "stage": "dd", "patterns": [r"\bsnda\b"]},
    {"key": "title", "label": "Title Commitment / Report", "stage": "dd", "patterns": [r"\btitle\b"]},
    {"key": "deposit", "label": "Deposit / Escrow / Wiring", "stage": "dd",
     "patterns": [r"deposit", r"wir(e|ing)", r"escrow"]},
    {"key": "reports", "label": "Third-Party Reports", "stage": "dd",
     "patterns": [r"phase\s*(1|i)\b", r"environmental", r"survey", r"\balta\b",
                  r"property\s+condition", r"\bpcr\b", r"appraisal", r"zoning", r"inspection"]},
    {"key": "settlement", "label": "Settlement Statement", "stage": "closing",
     "patterns": [r"settlement", r"closing\s+statement"]},
    {"key": "deed", "label": "Deed", "stage": "closing", "patterns": [r"\bdeed\b"]},
    {"key": "assignment", "label": "Assignment of Lease", "stage": "closing",
     "patterns": [r"assignment\s+of\s+lease", r"assignment"]},
    {"key": "notice_tenant", "label": "Notice to Tenant", "stage": "closing",
     "patterns": [r"notice\s+to\s+tenant"]},
    {"key": "binder", "label": "Closing Binder / Contact Sheet", "stage": "closing",
     "patterns": [r"closing\s+binder", r"contact\s+sheet", r"closing"]},
]
_COMPILED = [(it, [re.compile(p, re.I) for p in it["patterns"]]) for it in ITEMS]

# Every property folder is a copy of the brokerage "Continuum" template tree:
# blank MREIS forms, checklists, underwriting templates, and design assets that
# exist regardless of deal progress. They must never count as evidence.
NOISE_FOLDERS = {"agent information", "links", "document fonts", "images & resources",
                 "images and resources", "new images and resources", "headshot links"}
NOISE_NAME = re.compile(
    r"template|fillable|checklist|mreis|sale comp form|brokerage services|argus|"
    r"trade record|affidavit|placeholder|cdt_property|_property_state_city", re.I)
NOISE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".tif", ".tiff", ".ai", ".psd", ".indd",
              ".idml", ".eps", ".svg", ".lst", ".otf", ".ttf", ".html", ".sf", ".mp4", ".mov"}
_OM_FOLDER = re.compile(r"^(om|\d+ - om|.*-om folder)$", re.I)


def deal_files(files: list[dict]) -> list[dict]:
    """Drop template/asset noise; annotate each survivor with basename + folder hints."""
    out = []
    for f in files:
        parts = f["n"].split("\\")
        base = parts[-1]
        ext = ("." + base.rsplit(".", 1)[-1].lower()) if "." in base else ""
        if ext in NOISE_EXTS or NOISE_NAME.search(base):
            continue
        if any(seg.lower() in NOISE_FOLDERS for seg in parts[:-1]):
            continue
        in_om = any(_OM_FOLDER.match(seg) for seg in parts[:-1]) and ext == ".pdf"
        out.append({**f, "base": base, "in_om": in_om})
    return out


def _load(path, default):
    try:
        return json.loads(path.read_text())
    except (OSError, json.JSONDecodeError):
        return default


def _save(path, data):
    STATUS_DIR.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=1))


def list_properties() -> list[dict]:
    """All property folders: Box root (minus non-properties) + inside 0-* categories."""
    root = get_box_root()
    props = []
    root_items = list_children(root)
    for item in root_items:
        if item["is_dir"] and is_property(item["name"]):
            props.append({
                "name": item["name"],
                "box_path": f"{root}\\{item['name']}",
                "category": category_of(item["name"]),
            })
    for cat in CATEGORY_FOLDERS:
        if not any(i["name"] == cat and i["is_dir"] for i in root_items):
            continue
        for item in list_children(f"{root}\\{cat}"):
            if item["is_dir"]:
                props.append({
                    "name": item["name"],
                    "box_path": f"{root}\\{cat}\\{item['name']}",
                    "category": category_of(item["name"], parent_name=cat),
                })
    return props


def classify(files: list[dict]) -> dict:
    """Match a property's file listing against checklist items and derive its stage."""
    items_found: dict[str, list[str]] = {}
    for f in deal_files(files):
        base = f["base"]
        for item, patterns in _COMPILED:
            # match the file name only — template folder names like "Phase 5 - Closing"
            # exist in every property and say nothing about progress
            if any(p.search(base) for p in patterns) or (item["key"] == "om" and f["in_om"]):
                items_found.setdefault(item["key"], []).append(f["n"])

    stage_idx = 0
    stages_hit = {it["stage"] for it, _ in _COMPILED if it["key"] in items_found}
    for i, key in enumerate(STAGE_KEYS, start=1):
        if key in stages_hit:
            stage_idx = i

    mtimes = [f["m"] for f in files if f.get("m")]
    return {
        "auto_stage_idx": stage_idx,  # 0 = no evidence
        "items_found": items_found,
        "n_files": len(files),
        "last_activity": max(mtimes) if mtimes else None,
    }


def missing_items(items_found: dict, stage_idx: int) -> list[str]:
    """Item keys expected by the current stage that have no evidence."""
    if stage_idx <= 0:
        return []
    reached = set(STAGE_KEYS[:stage_idx])
    return [it["key"] for it in ITEMS
            if it["stage"] in reached and it["key"] not in items_found]


def run_scan(job_id: str) -> dict:
    """Scan every property folder (names only) in one PowerShell process."""
    props = list_properties()
    jobs.update(job_id, phase="scan", total=len(props), done=0)

    proc = subprocess.Popen(
        _ps_args("scan_tree.ps1"),
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, start_new_session=True,
    )
    by_path = {p["box_path"]: p for p in props}

    def feed():
        for p in props:
            proc.stdin.write(_b64(p["box_path"]) + "\n")
            proc.stdin.flush()
        proc.stdin.close()

    threading.Thread(target=feed, daemon=True).start()

    scan = _load(SCAN_PATH, {})
    done, errors = 0, 0
    for line in proc.stdout:
        line = line.strip().lstrip("﻿")
        if not line:
            continue
        try:
            result = json.loads(line)
        except json.JSONDecodeError:
            continue
        path = result.get("path", "")
        prop = by_path.get(path)
        if not prop:
            continue
        done += 1
        jobs.update(job_id, done=done, current_file=prop["name"])
        entry = {"name": prop["name"], "category": prop["category"], "scanned_at": time.time()}
        if "error" in result:
            errors += 1
            entry["error"] = result["error"][:200]
        else:
            # keep the raw name list so rules can evolve without a re-scan
            entry["files"] = [{"n": f["n"], "m": f.get("m")} for f in (result.get("files") or [])]
        scan[path] = entry
        if done % 5 == 0:  # checkpoint so the dashboard fills in as the scan runs
            with _lock:
                _save(SCAN_PATH, scan)
    proc.wait(timeout=60)

    scan = {p: v for p, v in scan.items() if p in by_path}
    with _lock:
        _save(SCAN_PATH, scan)
    return {"scanned": done, "errors": errors}


def start_scan() -> str:
    job_id = jobs.create("scan")

    def work():
        result = run_scan(job_id)
        jobs.update(job_id, message=f"Scanned {result['scanned']} properties "
                                    f"({result['errors']} errors).")

    jobs.run_in_thread(job_id, work)
    return job_id


def get_status() -> dict:
    """Everything the dashboard needs: per-property rows + stage/item metadata."""
    scan = _load(SCAN_PATH, {})
    overrides = _load(OVERRIDES_PATH, {})
    rows = []
    for path, entry in scan.items():
        if "files" in entry:
            info = classify(entry["files"])
        else:  # legacy scan entry (pre file-list format)
            info = {"auto_stage_idx": entry.get("auto_stage_idx", 0),
                    "items_found": {}, "n_files": entry.get("n_files", 0),
                    "last_activity": entry.get("last_activity")}
        override = overrides.get(path)
        auto_idx = info["auto_stage_idx"]
        idx = STAGE_KEYS.index(override) + 1 if override in STAGE_KEYS else auto_idx
        found = info["items_found"]
        rows.append({
            "box_path": path,
            "name": entry.get("name", path.split("\\")[-1]),
            "category": entry.get("category", "Other"),
            "stage_idx": idx,
            "stage": STAGE_KEYS[idx - 1] if idx else None,
            "auto_stage_idx": auto_idx,
            "override": override,
            "items_found": {k: v[:5] for k, v in found.items()},
            "missing": missing_items(found, idx),
            "n_files": info["n_files"],
            "last_activity": info["last_activity"],
            "error": entry.get("error"),
        })
    rows.sort(key=lambda r: (-r["stage_idx"], r["name"].lower()))
    scanned_at = max((e.get("scanned_at", 0) for e in scan.values()), default=None)
    return {
        "stages": [{"key": s["key"], "label": s["label"]} for s in STAGES],
        "items": [{"key": i["key"], "label": i["label"], "stage": i["stage"]} for i in ITEMS],
        "properties": rows,
        "scanned_at": scanned_at,
    }


def get_scan_entry(box_path: str) -> dict | None:
    return _load(SCAN_PATH, {}).get(box_path)


def set_override(box_path: str, stage: str | None) -> None:
    if stage is not None and stage not in STAGE_KEYS:
        raise ValueError(f"unknown stage {stage!r}")
    with _lock:
        overrides = _load(OVERRIDES_PATH, {})
        if stage is None:
            overrides.pop(box_path, None)
        else:
            overrides[box_path] = stage
        _save(OVERRIDES_PATH, overrides)
