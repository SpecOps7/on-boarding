"""Outlook hook (no API): pull mail from classic Outlook via COM, categorize each
message like the Box folders (property + document type, best judgement, else
leave alone), file it into Outlook subfolders, and turn workflow-relevant mail
into PROPOSALS that a human approves (go) or rejects (no-go).
"""

import json
import re
import subprocess
import threading
import time
import uuid
from datetime import datetime, timedelta
from pathlib import Path

from . import dates, jobs, status
from .box_bridge import PROJECT_DIR, _b64, _ps_args, wsl_to_win
from .categories import squash

MAIL_DIR = PROJECT_DIR / ".cache" / "outlook"
MAIL_PATH = MAIL_DIR / "mail.jsonl"
STATE_PATH = MAIL_DIR / "state.json"
PROPOSALS_PATH = MAIL_DIR / "proposals.json"
LEARNED_PATH = MAIL_DIR / "learned.json"
ATTACH_DIR = MAIL_DIR / "attachments"

_lock = threading.Lock()
_poller_started = False

DEFAULT_SETTINGS = {"auto_poll_minutes": 0, "since_days": 30, "overlap_hours": 48,
                    "file_into_folders": True, "folder_root": "Deals", "organize_bulk": True}
FOLDER_UNASSIGNED = "_Unassigned deal mail"
FOLDER_LISTINGS = "Listings & Marketing Blasts"
FOLDER_NOTIFICATIONS = "Notifications"

DEAL_DOC_EXTS = {".pdf", ".docx", ".doc", ".xlsx", ".xls"}
DATE_DOC_TYPES = {"psa", "loi", "settlement", "estoppel"}
DATE_WORDS = re.compile(r"closing|inspection|due diligence|\bdd\b|deposit|extend|extension|estoppel|effective date", re.I)
DRAFT_RE = re.compile(r"\bdraft\b|redline|for (your )?review|proposed|\bv\d+\b", re.I)
EXEC_RE = re.compile(r"fully[- ]executed|counter-?signed|\bsigned\b|\bexecuted\b|\bfe\b", re.I)

# Bulk / marketing detection — these never get matched to a property
BULK_SENDER = re.compile(
    r"no-?reply|newsletter|marketing|campaign|notification|listings?@|^emails?@|^info@|comms\.|"
    r"digital\.|crexi|rcm1\.com|ccsend\.com|savills\.info|francemedia|constantcontact|mailchimp|"
    r"hubspot|sendgrid|linkedin|dallasnews|ringcentral|@e\.|@em\.|@mail\.", re.I)
BLAST_SUBJECT = re.compile(
    r"^(just listed|new listing|new to market|for sale|available:|now available|price (reduction|reduced)|"
    r"new price|featured|auction|under contract|just sold|closed:|webinar|bulletin|newsletter)|"
    r"recommended for you|\|.*\|.*\||\bcap\b.*\|", re.I)
NOTIFICATION_SENDER = re.compile(r"ringcentral|linkedin|no-?reply|notification|donotreply|alerts?@|calendar", re.I)

STREET_STOP = {"st", "ave", "rd", "blvd", "dr", "hwy", "pkwy", "ln", "ct", "cir", "way", "street",
               "avenue", "road", "boulevard", "drive", "n", "s", "e", "w", "ne", "nw", "se", "sw",
               "north", "south", "east", "west", "us", "state", "route", "terr", "pike"}
# Words too common to identify a property on their own
COMMON = {"green", "valley", "ranch", "square", "main", "park", "village", "center", "centre",
          "plaza", "lake", "river", "hill", "hills", "point", "creek", "spring", "springs", "grove",
          "oak", "pine", "cedar", "maple", "university", "county", "highway", "medical", "dental",
          "urgent", "care", "health", "family", "general", "dollar", "store", "commons", "crossing",
          "market", "station", "town", "city", "first", "second", "old", "new", "texas", "indiana",
          "florida", "ohio", "kentucky", "oklahoma", "tennessee", "michigan", "arizona", "avenue",
          "terrace", "circle", "washington", "madison", "santa", "school", "phoenix", "memorial"}


# ---------------------------------------------------------------- persistence
def _load_json(path: Path, default):
    try:
        return json.loads(path.read_text())
    except (OSError, json.JSONDecodeError):
        return default


def _save_json(path: Path, data) -> None:
    MAIL_DIR.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=1))


def get_settings() -> dict:
    return {**DEFAULT_SETTINGS, **_load_json(STATE_PATH, {}).get("settings", {})}


def set_settings(new: dict) -> dict:
    with _lock:
        state = _load_json(STATE_PATH, {})
        state["settings"] = {**get_settings(), **new}
        _save_json(STATE_PATH, state)
    return state["settings"]


def load_mail() -> dict[str, dict]:
    msgs = {}
    if MAIL_PATH.exists():
        for line in MAIL_PATH.read_text().splitlines():
            if line.strip():
                try:
                    m = json.loads(line)
                    msgs[m["entry_id"]] = m
                except json.JSONDecodeError:
                    continue
    return msgs


def _rewrite_mail(msgs: dict[str, dict]) -> None:
    MAIL_DIR.mkdir(parents=True, exist_ok=True)
    MAIL_PATH.write_text("".join(json.dumps(m) + "\n" for m in msgs.values()))


def _append_mail(new: list[dict]) -> None:
    MAIL_DIR.mkdir(parents=True, exist_ok=True)
    with MAIL_PATH.open("a") as f:
        for m in new:
            f.write(json.dumps(m) + "\n")


# ---------------------------------------------------------------- Outlook pull
def pull(job_id: str | None = None, since: datetime | None = None) -> list[dict]:
    """Run the COM bridge; return NEW messages (deduped by entry_id) and advance the watermark."""
    state = _load_json(STATE_PATH, {})
    settings = get_settings()
    if since is None:
        if state.get("watermark"):
            # overlap so mail that syncs in late (or arrives out of order) is not missed
            since = datetime.fromisoformat(state["watermark"]) - timedelta(hours=settings["overlap_hours"])
        else:
            since = datetime.now() - timedelta(days=settings["since_days"])
    ATTACH_DIR.mkdir(parents=True, exist_ok=True)
    proc = subprocess.Popen(
        _ps_args("outlook_pull.ps1", "-SinceB64", _b64(since.isoformat()),
                 "-AttachDirB64", _b64(wsl_to_win(ATTACH_DIR))),
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, start_new_session=True,
    )
    last = [time.monotonic()]
    stop = threading.Event()

    def watchdog():  # Outlook COM can hang on dialogs — don't let the job hang forever
        while not stop.wait(10):
            if time.monotonic() - last[0] > 240:
                proc.kill()
                return

    threading.Thread(target=watchdog, daemon=True).start()

    known = load_mail()
    known_ids = set(known)
    for m in known.values():
        if m.get("original_entry_id"):
            known_ids.add(m["original_entry_id"])
        known_ids.update(m.get("alias_ids") or [])
    new, seen, max_received, error = [], set(), None, None
    try:
        for line in proc.stdout:
            last[0] = time.monotonic()
            line = line.strip().lstrip("﻿")
            if not line:
                continue
            try:
                ev = json.loads(line)
            except json.JSONDecodeError:
                continue
            kind = ev.get("event")
            if kind == "mail":
                eid = ev["entry_id"]
                if eid in known_ids or eid in seen:
                    continue
                seen.add(eid)
                for a in ev.get("attachments") or []:
                    if a.get("saved_path"):
                        rel = a["saved_path"].split("attachments", 1)[-1].lstrip("\\").replace("\\", "/")
                        a["local_path"] = str(ATTACH_DIR / rel)
                ev.pop("event", None)
                new.append(ev)
                if job_id:
                    jobs.update(job_id, done=len(new), current_file=ev.get("subject", "")[:80])
            elif kind == "done":
                max_received = ev.get("max_received")
            elif kind == "error":
                error = ev.get("message")
        proc.wait(timeout=30)
    finally:
        stop.set()
    if error:
        raise RuntimeError(error)
    if max_received is None and proc.returncode not in (0, None):
        raise RuntimeError(f"outlook_pull failed: {proc.stderr.read()[:300]}")
    if max_received:
        with _lock:
            state = _load_json(STATE_PATH, {})
            state["watermark"] = max_received[:19]
            state["last_pull"] = time.time()
            _save_json(STATE_PATH, state)
    return new


# ---------------------------------------------------------------- property matching
def is_bulk(msg: dict) -> str | None:
    """'notification' | 'listing' | None — bulk mail is organized but never matched to a property."""
    sender = (msg.get("sender_email") or "").lower()
    subject = msg.get("subject") or ""
    if NOTIFICATION_SENDER.search(sender):
        return "notification"
    if BULK_SENDER.search(sender) or BLAST_SUBJECT.search(subject):
        return "listing"
    return None


class PropertyMatcher:
    """Best-judgement match of an email to a known property folder, like the Box categorization.

    Strong evidence: exact folder name, brand + street number, or the street number adjacent to
    the street name ("603 Hwy 35", "3811 Commons"). Body text only contributes those specific
    forms — brand words or common street words alone in a body never match (listing blasts
    mention dozens of addresses)."""

    def __init__(self, props: list[dict]):
        self.props = []
        for p in props:
            name = p["name"]
            head, rest = self._split(name)
            num = re.search(r"\b(\d{2,6})\b", rest)
            brand_words = set(re.findall(r"[a-z]+", head.lower()))
            street_words = [w for w in re.findall(r"[a-z]+", rest.lower())
                            if w not in STREET_STOP and w not in brand_words and len(w) > 2]
            self.props.append({
                **p, "key": squash(name),
                "brand": squash(head) if len(squash(head)) >= 3 else "",
                "brand_re": re.compile(r"\b" + r"\s*".join(map(re.escape, sorted(brand_words, key=len, reverse=True)[:2])) + r"\b", re.I)
                            if brand_words else None,
                "number": num.group(1) if num else "",
                "street_words": street_words[:4],
                "distinct_words": [w for w in street_words if len(w) >= 5 and w not in COMMON][:4],
            })

    @staticmethod
    def _split(name: str) -> tuple[str, str]:
        # buy-side process folders: "DD_Items_Fast Pace…", "Due Diligence_Nextcare - …", "Buyer's_DD_Emcura…"
        name = re.sub(r"^(buyer'?s?_?|dd_|due diligence(\s+items)?_|dd_items_)+", "", name, flags=re.I)
        name = re.sub(r"_(seller|buyer|due diligence folder)$", "", name, flags=re.I)
        parts = re.split(r"\s*-\s+|_", name, maxsplit=1)
        paren = " ".join(re.findall(r"\((.*?)\)", parts[0]))
        head = re.sub(r"\(.*?\)", " ", parts[0]).strip()
        rest = (paren + " " + (parts[1] if len(parts) > 1 else "")).strip()
        return head, rest

    def _phrase(self, p: dict, text: str) -> bool:
        """street number followed (within two tokens) by one of the street words."""
        if not p["number"] or not p["street_words"]:
            return False
        alt = "|".join(re.escape(w) for w in p["street_words"])
        return re.search(r"\b" + re.escape(p["number"]) + r"\b(?:\s+\S+){0,2}\s+(?:" + alt + r")\b", text, re.I) is not None

    def match(self, head_text: str, body_text: str, sender_email: str = "", learned: dict | None = None) -> dict | None:
        head_sq, body_sq = squash(head_text), squash(body_text)
        learned = learned or {}
        aliases = load_aliases()
        alias_hits = {a["box_path"] for a in aliases
                      if re.search(a["pattern"], head_text + " " + body_text[:3000], re.I)}
        scored = []
        for p in self.props:
            score, why = 0, []
            if p["key"] and (p["key"] in head_sq or p["key"] in body_sq):
                score += 100; why.append("exact folder name")
            if p["box_path"] in alias_hits:
                score += 50; why.append("alias")
            brand_head = bool(p["brand"]) and (p["brand"] in head_sq or
                                               (p["brand_re"] is not None and p["brand_re"].search(head_text) is not None))
            brand_any = brand_head or (bool(p["brand"]) and p["brand"] in body_sq)
            num_head = bool(p["number"]) and re.search(r"\b" + re.escape(p["number"]) + r"\b", head_text) is not None
            phrase = self._phrase(p, head_text) or self._phrase(p, body_text)
            if phrase:
                score += 55; why.append("street number + street name")
            if brand_head and num_head and len(p["number"]) >= 3:
                score += 60; why.append("brand + street number in subject")
            elif brand_any and phrase:
                score += 25; why.append("brand")
            elif brand_head and any(re.search(r"\b" + re.escape(w) + r"\b", head_text, re.I) for w in p["distinct_words"]):
                score += 45; why.append("brand + location in subject")
            elif brand_head:
                score += 10; why.append("brand only")
            dist_head = sum(1 for w in p["distinct_words"] if re.search(r"\b" + re.escape(w) + r"\b", head_text, re.I))
            if dist_head >= 2 and not phrase:
                score += 45; why.append("multiple location words in subject")
            if sender_email and learned.get(sender_email.lower()) == p["box_path"]:
                score += 30; why.append("known sender")
            if score:
                scored.append((score, p, why))
        if not scored:
            return None
        scored.sort(key=lambda t: -t[0])
        best, second = scored[0], (scored[1] if len(scored) > 1 else None)
        if best[0] < 45:
            return None  # weak → leave it alone
        if second and second[0] >= best[0] - 15 and "alias" not in best[2]:
            return None  # ambiguous between folders (unless an alias decided it) → leave it alone
        return {"box_path": best[1]["box_path"], "name": best[1]["name"],
                "category": best[1].get("category", "Other"),
                "confidence": min(0.99, best[0] / 100), "reason": ", ".join(best[2])}


ALIASES_PATH = MAIL_DIR / "aliases.json"
# Deal nicknames that don't appear in the Box folder name (editable in .cache/outlook/aliases.json).
# An alias adds a strong vote for one folder — it decides ties between sibling folders of one deal.
DEFAULT_ALIASES = [
    {"pattern": r"port\s+lavaca|603\s+hwy\s+35", "box_path": r"C:\Users\ctith\Box\NextCare - 603 Texas 35_2"},
    {"pattern": r"del\s+city", "box_path": r"C:\Users\ctith\Box\0-Urgent Care\Access Medical Center - Del City OK"},
    {"pattern": r"frankfort|301\s+versailles", "box_path": r"C:\Users\ctith\Box\Due Diligence_Fast Pace_301 Versailles Rd_Frankfort KY"},
    {"pattern": r"yukon", "box_path": r"C:\Users\ctith\Box\DD_Access Medical Center_Yukon OK"},
    {"pattern": r"grosse\s+pointe|emcura", "box_path": r"C:\Users\ctith\Box\DD_Emcura Immediate & Primary Care_Grosse Pointe Woods MI"},
    {"pattern": r"lafayette", "box_path": r"C:\Users\ctith\Box\DD_Items_Fast Pace Urgent Care_Lafayette TN"},
]


def load_aliases() -> list[dict]:
    aliases = _load_json(ALIASES_PATH, None)
    if aliases is None:
        aliases = DEFAULT_ALIASES
        _save_json(ALIASES_PATH, aliases)
    return aliases


def _properties() -> list[dict]:
    scan = status._load(status.SCAN_PATH, {})
    if scan:
        return [{"box_path": p, "name": e.get("name", p.split("\\")[-1]),
                 "category": e.get("category", "Other")} for p, e in scan.items()]
    return status.list_properties()


# ---------------------------------------------------------------- doc classification
GENERIC_IN_SUBJECT = {r"closing", r"\bcontract\b", r"\btitle\b", r"\blease\b", r"inspection",
                      r"financial", r"marketing", r"abstract", r"assignment", r"\blisting\b",
                      r"survey", r"deposit", r"insurance", r"zoning", r"\bom\b", r"blast", r"flyer"}


def classify_doc(msg: dict) -> dict | None:
    atts = [(a.get("name") or "") for a in msg.get("attachments") or []]
    subject = msg.get("subject") or ""
    hits = {}
    for item, patterns in status._COMPILED:
        if any(p.search(a) for p in patterns for a in atts):
            hits[item["key"]] = ("attachment", item)
        elif any(p.search(subject) for p in patterns if p.pattern not in GENERIC_IN_SUBJECT):
            hits[item["key"]] = ("subject", item)
    if not hits:
        return None
    key, (source, item) = max(hits.items(), key=lambda kv: (status.STAGE_KEYS.index(kv[1][1]["stage"]),
                                                           kv[1][0] == "attachment"))
    signal_text = subject + " " + " ".join(atts)
    return {"item": key, "label": item["label"], "stage": item["stage"],
            "implied_stage_idx": status.STAGE_KEYS.index(item["stage"]) + 1,
            "source": source, "strong": source == "attachment",
            "executed": bool(EXEC_RE.search(signal_text)), "draft": bool(DRAFT_RE.search(signal_text))}


def _extract_dates_from_mail(msg: dict) -> dict | None:
    key = msg["entry_id"][-16:]
    d = ATTACH_DIR / key
    d.mkdir(parents=True, exist_ok=True)
    body_txt = d / "email_body.txt"
    body_txt.write_text(f"Subject: {msg.get('subject','')}\nFrom: {msg.get('sender_name','')} "
                        f"<{msg.get('sender_email','')}>\nDate: {msg.get('received','')}\n\n{msg.get('body','')}")
    paths = [str(body_txt)]
    for a in (msg.get("attachments") or [])[:4]:
        lp = a.get("local_path")
        if lp and Path(lp).exists() and Path(lp).suffix.lower() in {".pdf", ".docx", ".txt"}:
            r = dates._readable_copy(lp)
            if r:
                paths.append(r)
    try:
        return dates._extract_with_claude(paths)
    except Exception as e:  # noqa: BLE001
        return {"error": str(e)[:200]}


# ---------------------------------------------------------------- Outlook filing
def _safe_folder(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "-", name).strip()


def target_folder(m: dict, settings: dict | None = None) -> str | None:
    """Where this message belongs (relative to Inbox); '' = Inbox root; None = don't touch."""
    settings = settings or get_settings()
    if m.get("folder") != "Inbox":
        return None
    root = settings.get("folder_root") or "Deals"
    if m.get("match"):
        return "\\".join([root, _safe_folder(m["match"].get("category") or "Other"), _safe_folder(m["match"]["name"])])
    if settings.get("organize_bulk"):
        if m.get("bulk") == "notification":
            return FOLDER_NOTIFICATIONS
        if m.get("bulk") == "listing":
            return FOLDER_LISTINGS
        if m.get("doc"):  # deal-looking mail with no known property
            return root + "\\" + FOLDER_UNASSIGNED
    if m.get("outlook_folder"):  # previously filed, no longer categorized → back to the Inbox
        return ""
    return None


def file_into_outlook(msgs: list[dict], job_id: str | None = None) -> dict:
    """Move messages to their target folders (created on demand). Moving changes Outlook
    EntryIDs, so message records are updated in place."""
    settings = get_settings()
    if not settings.get("file_into_folders"):
        return {"moved": 0, "errors": 0}
    todo = []
    for m in msgs:
        t = target_folder(m, settings)
        if t is None:
            continue
        current = (m.get("outlook_folder") or "").split("\\Inbox", 1)[-1].strip("\\")
        if m.get("outlook_folder") and current.lower() == t.lower():
            continue
        todo.append((m, t))
    if not todo:
        return {"moved": 0, "errors": 0}
    if job_id:
        jobs.update(job_id, phase="file", total=len(todo), done=0, current_file="")
    by_id = {m["entry_id"]: m for m, _ in todo}
    lines = [m["entry_id"] + "|" + (_b64(t) if t else "") for m, t in todo]
    try:
        proc = subprocess.run(_ps_args("outlook_move.ps1"), input="\n".join(lines) + "\n",
                              capture_output=True, text=True, timeout=90 + 3 * len(todo))
        out = proc.stdout
    except subprocess.TimeoutExpired as e:
        out = (e.stdout.decode() if isinstance(e.stdout, bytes) else e.stdout) or ""
        for m, _ in todo:
            m["file_error"] = "Outlook did not respond (busy or showing a dialog) — will retry on next check"
    moved, errors = 0, 0
    for i, line in enumerate(out.splitlines(), 1):
        line = line.strip().lstrip("﻿")
        if not line:
            continue
        try:
            r = json.loads(line)
        except json.JSONDecodeError:
            continue
        m = by_id.get(r.get("entry_id"))
        if not m:
            continue
        if job_id:
            jobs.update(job_id, done=i, current_file=(m.get("subject") or "")[:80])
        if r.get("ok"):
            m["outlook_folder"] = r.get("folder")
            m.pop("file_error", None)
            if r.get("new_entry_id") and r["new_entry_id"] != m["entry_id"]:
                m.setdefault("original_entry_id", m["entry_id"])
                m["entry_id"] = r["new_entry_id"]
            moved += 0 if r.get("skipped") else 1
        else:
            errors += 1
            m["file_error"] = (r.get("error") or "")[:200]
    return {"moved": moved, "errors": errors}


# ---------------------------------------------------------------- proposals
def _current_rows() -> dict[str, dict]:
    return {r["box_path"]: r for r in status.get_status()["properties"]}


def _stage_label(idx: int) -> str:
    return "Not started" if idx <= 0 else status.STAGES[idx - 1]["label"]


def build_changes(msg: dict, match: dict, doc: dict | None, extracted: dict | None, rows: dict) -> list[dict]:
    box_path = match["box_path"]
    row = rows.get(box_path, {"stage_idx": 0})
    scan_entry = status.get_scan_entry(box_path) or {}
    existing_files = {f["n"].split("\\")[-1].lower() for f in scan_entry.get("files", [])}
    changes = []
    if doc and doc["strong"] and not doc["draft"] and doc["implied_stage_idx"] > row["stage_idx"]:
        changes.append({"kind": "stage", "from": row["stage_idx"], "to": doc["implied_stage_idx"],
                        "label": f"{_stage_label(row['stage_idx'])} → {_stage_label(doc['implied_stage_idx'])}"})
    if extracted and not extracted.get("error"):
        cur = dates._load().get(box_path, {})
        for field, label in dates.DATE_FIELDS:
            new = extracted.get(field)
            if new and re.match(r"^\d{4}-\d{2}-\d{2}$", str(new)) and new != cur.get(field):
                changes.append({"kind": "date", "field": field, "from": cur.get(field), "to": new,
                                "label": f"{label}: {cur.get(field) or '—'} → {new}"})
    for a in msg.get("attachments") or []:
        name, lp = a.get("name") or "", a.get("local_path")
        if not lp or Path(name).suffix.lower() not in DEAL_DOC_EXTS or name.lower() in existing_files:
            continue
        sub = doc["label"] if doc else "Email"
        changes.append({"kind": "file", "attachment": name, "local_path": lp,
                        "target": f"{box_path}\\Inbox Filing\\{_safe_folder(sub)}\\{name}",
                        "label": f"File {name} → {match['name']}\\Inbox Filing\\{sub}"})
    return changes


def _change_key(c: dict) -> tuple:
    return (c["kind"], c.get("to"), c.get("field"), c.get("attachment"))


def upsert_proposal(proposals: list[dict], msg: dict, match: dict, doc: dict | None, changes: list[dict]) -> bool:
    """Merge into a pending proposal for the same thread + property, else create one. Returns True if new."""
    if not changes:
        return False
    email = {"subject": msg.get("subject"), "sender": msg.get("sender_name"),
             "sender_email": msg.get("sender_email"), "received": msg.get("received"),
             "folder": msg.get("folder"), "excerpt": (msg.get("body") or "")[:700],
             "attachments": [a.get("name") for a in msg.get("attachments") or []]}
    for p in proposals:
        if p["status"] == "pending" and p["box_path"] == match["box_path"] and \
                p.get("conversation_id") == msg.get("conversation_id"):
            known = {_change_key(c) for c in p["changes"]}
            added = [c for c in changes if _change_key(c) not in known]
            if added:
                p["changes"].extend(added)
            if (msg.get("received") or "") > (p["email"].get("received") or ""):
                p["email"] = email
                p["entry_id"] = msg["entry_id"]
            return False
    proposals.append({
        "id": uuid.uuid4().hex[:10], "entry_id": msg["entry_id"], "conversation_id": msg.get("conversation_id"),
        "status": "pending", "created": time.time(), "box_path": match["box_path"],
        "property_name": match["name"], "confidence": match["confidence"], "match_reason": match["reason"],
        "doc_type": doc["label"] if doc else None, "executed": bool(doc and doc["executed"]),
        "draft": bool(doc and doc["draft"]), "email": email,
        "notes": (msg.get("extracted") or {}).get("notes") if msg.get("extracted") and not msg["extracted"].get("error") else None,
        "changes": changes,
    })
    return True


# ---------------------------------------------------------------- categorize
def _match_one(matcher: PropertyMatcher, m: dict, learned: dict, thread_prop: dict) -> dict | None:
    if m.get("bulk"):
        return None
    head = " ".join([m.get("subject") or "", " ".join(a.get("name") or "" for a in m.get("attachments") or [])])
    match = matcher.match(head, (m.get("body") or "")[:6000], m.get("sender_email") or "", learned)
    if not match and m.get("conversation_id") in thread_prop:
        match = {**thread_prop[m["conversation_id"]], "confidence": 0.5, "reason": "same email thread"}
    return match


def categorize(msgs: list[dict], job_id: str | None = None, with_dates: bool = True,
               file_mail: bool = True) -> dict:
    """Match + classify, file into Outlook subfolders, then propose. Runs date extraction at
    most once per thread per run (newest message)."""
    matcher = PropertyMatcher(_properties())
    learned = _load_json(LEARNED_PATH, {})
    rows = _current_rows()
    all_mail = load_mail()
    thread_prop = {m.get("conversation_id"): m["match"] for m in all_mail.values()
                   if m.get("match") and m.get("conversation_id")}
    matched = 0
    for i, m in enumerate(msgs, 1):
        if job_id:
            jobs.update(job_id, phase="categorize", done=i, total=len(msgs), current_file=(m.get("subject") or "")[:80])
        m["bulk"] = is_bulk(m)
        m["match"] = _match_one(matcher, m, learned, thread_prop)
        m["doc"] = classify_doc(m)
        if m["match"]:
            matched += 1
            thread_prop.setdefault(m.get("conversation_id"), m["match"])

    filed = file_into_outlook(msgs, job_id) if file_mail else {"moved": 0, "errors": 0}

    proposals = _load_json(PROPOSALS_PATH, [])
    made, extracted_threads = 0, set()
    ordered = sorted([m for m in msgs if m.get("match")], key=lambda m: m.get("received") or "", reverse=True)
    for i, m in enumerate(ordered, 1):
        match, doc = m["match"], m.get("doc")
        if job_id:
            jobs.update(job_id, phase="propose", done=i, total=len(ordered), current_file=(m.get("subject") or "")[:80])
        extracted = None
        wants = with_dates and ((doc and doc["item"] in DATE_DOC_TYPES) or
                                (DATE_WORDS.search(m.get("body") or "") and m.get("attachments")))
        conv = m.get("conversation_id") or m["entry_id"]
        if wants and conv not in extracted_threads:
            extracted_threads.add(conv)
            extracted = _extract_dates_from_mail(m)
            m["extracted"] = extracted
        changes = build_changes(m, match, doc, extracted, rows)
        if upsert_proposal(proposals, m, match, doc, changes):
            made += 1
    with _lock:
        _save_json(PROPOSALS_PATH, proposals)
    return {"messages": len(msgs), "matched": matched, "proposals": made, **filed}


def dedupe_stored() -> dict:
    """Collapse records that are the same physical email under different Outlook EntryIDs.

    During the initial mailbox sync, Exchange can re-deliver an Inbox copy of a message we
    already moved — the old record's EntryID goes dead ("message cannot be found") and the
    re-delivered copy gets pulled as new. Keep the best record per (subject, received, sender)
    and remember every collapsed id so future pulls skip them."""
    msgs = load_mail()
    groups: dict[tuple, list[dict]] = {}
    for m in msgs.values():
        key = (m.get("subject") or "", (m.get("received") or "")[:19], (m.get("sender_email") or "").lower())
        groups.setdefault(key, []).append(m)

    def rank(m: dict) -> tuple:
        return (bool(m.get("outlook_folder")) and not m.get("file_error"),
                bool(m.get("match")), not m.get("file_error"), m.get("received") or "")

    kept: dict[str, dict] = {}
    removed = 0
    for group in groups.values():
        group.sort(key=rank, reverse=True)
        best = group[0]
        aliases = set(best.get("alias_ids") or [])
        for other in group[1:]:
            removed += 1
            aliases.add(other["entry_id"])
            if other.get("original_entry_id"):
                aliases.add(other["original_entry_id"])
            if not best.get("match") and other.get("match"):
                best["match"] = other["match"]
        if aliases:
            best["alias_ids"] = sorted(aliases)
        if best.get("file_error") and any(o.get("outlook_folder") for o in group[1:]):
            best.pop("file_error", None)
        kept[best["entry_id"]] = best
    _rewrite_mail(kept)

    # retire pending proposals that point at removed records
    alive = set(kept)
    for m in kept.values():
        alive.update(m.get("alias_ids") or [])
        if m.get("original_entry_id"):
            alive.add(m["original_entry_id"])
    with _lock:
        props = _load_json(PROPOSALS_PATH, [])
        retired = 0
        for p in props:
            if p["status"] == "pending" and p["entry_id"] not in alive:
                p["status"] = "superseded"
                retired += 1
        _save_json(PROPOSALS_PATH, props)
    return {"kept": len(kept), "removed": removed, "proposals_retired": retired}


def refile_pending() -> dict:
    """Retry filing for stored mail whose move failed or whose target changed."""
    msgs = load_mail()
    todo = [m for m in msgs.values() if target_folder(m) is not None]
    r = file_into_outlook(todo)
    _rewrite_mail({m["entry_id"]: m for m in msgs.values()})
    return r


def recategorize_all(job_id: str | None = None, with_dates: bool = False) -> dict:
    """Re-run matching/classification over every stored message with the current rules,
    re-file whatever changed, retire pending proposals whose match changed, and rebuild them
    (reusing already-extracted dates — no new Claude calls unless with_dates)."""
    msgs = load_mail()
    ordered = sorted(msgs.values(), key=lambda m: m.get("received") or "")
    matcher = PropertyMatcher(_properties())
    learned = _load_json(LEARNED_PATH, {})
    thread_prop: dict = {}
    changed, matched = 0, 0
    for i, m in enumerate(ordered, 1):
        if job_id:
            jobs.update(job_id, phase="categorize", done=i, total=len(ordered), current_file=(m.get("subject") or "")[:80])
        old = (m.get("match") or {}).get("box_path")
        m["bulk"] = is_bulk(m)
        m["match"] = _match_one(matcher, m, learned, thread_prop)
        m["doc"] = classify_doc(m)
        if m["match"]:
            matched += 1
            thread_prop.setdefault(m.get("conversation_id"), m["match"])
        if (m.get("match") or {}).get("box_path") != old:
            changed += 1
    filed = file_into_outlook(ordered, job_id)
    _rewrite_mail({m["entry_id"]: m for m in ordered})

    proposals = _load_json(PROPOSALS_PATH, [])
    by_entry = {}
    for m in ordered:
        by_entry[m["entry_id"]] = m
        if m.get("original_entry_id"):
            by_entry[m["original_entry_id"]] = m
    for p in proposals:
        m = by_entry.get(p["entry_id"])
        if p["status"] == "pending" and (not m or (m.get("match") or {}).get("box_path") != p["box_path"]):
            p["status"] = "superseded"
    rows = _current_rows()
    made = 0
    for m in sorted([m for m in ordered if m.get("match")], key=lambda m: m.get("received") or "", reverse=True):
        extracted = m.get("extracted")
        if extracted is None and with_dates:
            doc = m.get("doc")
            if doc and doc["item"] in DATE_DOC_TYPES:
                extracted = _extract_dates_from_mail(m)
                m["extracted"] = extracted
        changes = build_changes(m, m["match"], m.get("doc"), extracted, rows)
        if upsert_proposal(proposals, m, m["match"], m.get("doc"), changes):
            made += 1
    with _lock:
        _save_json(PROPOSALS_PATH, proposals)
    _rewrite_mail({m["entry_id"]: m for m in ordered})
    return {"messages": len(ordered), "matched": matched, "changed": changed, "proposals": made, **filed}


def run_pull_job(job_id: str) -> dict:
    jobs.update(job_id, phase="pull", message="Pulling mail from classic Outlook…", done=0, total=0)
    new = pull(job_id)
    result = categorize(new, job_id)
    _append_mail(new)
    dedupe_stored()
    retry = refile_pending()
    result["moved"] += retry["moved"]
    return {"new": len(new), **result}


def start_pull() -> str:
    job_id = jobs.create("mail")

    def work():
        r = run_pull_job(job_id)
        jobs.update(job_id, message=f"{r['new']} new emails · {r['matched']} matched to properties · "
                                    f"{r['moved']} filed into Outlook folders · {r['proposals']} proposals to review.")

    jobs.run_in_thread(job_id, work)
    return job_id


def start_recategorize() -> str:
    job_id = jobs.create("mail-recategorize")

    def work():
        r = recategorize_all(job_id)
        jobs.update(job_id, message=f"Re-categorized {r['messages']} emails · {r['matched']} matched · "
                                    f"{r['changed']} changed · {r['moved']} re-filed · {r['proposals']} new proposals.")

    jobs.run_in_thread(job_id, work)
    return job_id


# ---------------------------------------------------------------- decisions (go / no-go)
def list_proposals(only: str = "pending") -> list[dict]:
    props = _load_json(PROPOSALS_PATH, [])
    if only != "all":
        props = [p for p in props if p["status"] == only]
    return sorted(props, key=lambda p: -p["created"])


def _apply_change(prop: dict, ch: dict) -> dict:
    try:
        if ch["kind"] == "stage":
            status.set_override(prop["box_path"], status.STAGE_KEYS[ch["to"] - 1])
        elif ch["kind"] == "date":
            d = dates._load()
            entry = d.setdefault(prop["box_path"], {"name": prop["property_name"]})
            entry[ch["field"]] = ch["to"]
            entry.setdefault("locked", [])
            if ch["field"] not in entry["locked"]:
                entry["locked"].append(ch["field"])
            entry.setdefault("sources", []).append(
                f"email: {prop['email'].get('subject')} ({(prop['email'].get('received') or '')[:10]})")
            dates._save(d)
        elif ch["kind"] == "file":
            proc = subprocess.run(
                _ps_args("copy_to_box.ps1", "-SrcB64", _b64(wsl_to_win(ch["local_path"])),
                         "-DstB64", _b64(ch["target"])),
                capture_output=True, text=True, timeout=300)
            out = proc.stdout.strip().lstrip("﻿")
            r = json.loads(out) if out else {"ok": False, "error": proc.stderr[:200]}
            if not r.get("ok"):
                return {"ok": False, "error": r.get("error", "copy failed")}
        return {"ok": True}
    except Exception as e:  # noqa: BLE001
        return {"ok": False, "error": f"{type(e).__name__}: {e}"[:200]}


def decide(proposal_id: str, decision: str, accept: list[int] | None = None) -> dict:
    with _lock:
        props = _load_json(PROPOSALS_PATH, [])
        prop = next((p for p in props if p["id"] == proposal_id), None)
        if not prop:
            raise KeyError(proposal_id)
        if prop["status"] != "pending":
            return prop
        if decision == "reject":
            prop["status"] = "rejected"
        else:
            idxs = set(range(len(prop["changes"]))) if accept is None else set(accept)
            for i, ch in enumerate(prop["changes"]):
                ch["accepted"] = i in idxs
                ch["result"] = _apply_change(prop, ch) if i in idxs else {"ok": None}
            prop["status"] = "approved"
            learned = _load_json(LEARNED_PATH, {})
            if prop["email"].get("sender_email") and not BULK_SENDER.search(prop["email"]["sender_email"]):
                learned[prop["email"]["sender_email"].lower()] = prop["box_path"]
                _save_json(LEARNED_PATH, learned)
        prop["decided_at"] = time.time()
        _save_json(PROPOSALS_PATH, props)
        return prop


def assign(entry_id: str, box_path: str) -> dict | None:
    """Manually assign an email to a property; re-files it and rebuilds its proposal."""
    msgs = load_mail()
    m = msgs.get(entry_id) or next((x for x in msgs.values() if x.get("original_entry_id") == entry_id), None)
    if not m:
        raise KeyError(entry_id)
    prop_info = next((p for p in _properties() if p["box_path"] == box_path), None)
    if not prop_info:
        raise KeyError(box_path)
    m["match"] = {**prop_info, "confidence": 1.0, "reason": "assigned manually"}
    m["bulk"] = None
    m["doc"] = classify_doc(m)
    if m.get("extracted") is None and m["doc"] and m["doc"]["item"] in DATE_DOC_TYPES:
        m["extracted"] = _extract_dates_from_mail(m)
    old_id = m["entry_id"]
    file_into_outlook([m])
    if m["entry_id"] != old_id:
        msgs.pop(old_id, None)
    msgs[m["entry_id"]] = m
    _rewrite_mail(msgs)
    with _lock:
        props = _load_json(PROPOSALS_PATH, [])
        for p in props:
            if p["entry_id"] in (old_id, m["entry_id"]) and p["status"] == "pending":
                p["status"] = "superseded"
        changes = build_changes(m, m["match"], m["doc"], m.get("extracted"), _current_rows())
        new = None
        if upsert_proposal(props, m, m["match"], m["doc"], changes):
            new = props[-1]
        _save_json(PROPOSALS_PATH, props)
    learned = _load_json(LEARNED_PATH, {})
    se = (m.get("sender_email") or "").lower()
    if se and not BULK_SENDER.search(se):
        learned[se] = box_path
        _save_json(LEARNED_PATH, learned)
    return new


def list_unmatched(limit: int = 150) -> list[dict]:
    msgs = [m for m in load_mail().values() if not m.get("match") and not m.get("bulk")]
    msgs.sort(key=lambda m: m.get("received") or "", reverse=True)
    return [{"entry_id": m["entry_id"], "subject": m.get("subject"), "sender": m.get("sender_name"),
             "sender_email": m.get("sender_email"), "received": m.get("received"),
             "attachments": [a.get("name") for a in m.get("attachments") or []],
             "doc_type": (m.get("doc") or {}).get("label"), "outlook_folder": m.get("outlook_folder")}
            for m in msgs[:limit]]


def list_filed(limit: int = 300) -> list[dict]:
    msgs = [m for m in load_mail().values() if m.get("match") or m.get("bulk")]
    msgs.sort(key=lambda m: m.get("received") or "", reverse=True)
    return [{"entry_id": m["entry_id"], "subject": m.get("subject"), "sender": m.get("sender_name"),
             "received": m.get("received"),
             "property": m["match"]["name"] if m.get("match") else ("(" + m["bulk"] + ")"),
             "category": (m.get("match") or {}).get("category"), "confidence": (m.get("match") or {}).get("confidence"),
             "reason": (m.get("match") or {}).get("reason") or m.get("bulk"),
             "doc_type": (m.get("doc") or {}).get("label"), "outlook_folder": m.get("outlook_folder"),
             "file_error": m.get("file_error"), "folder": m.get("folder")} for m in msgs[:limit]]


def summary() -> dict:
    msgs = load_mail()
    props = _load_json(PROPOSALS_PATH, [])
    state = _load_json(STATE_PATH, {})
    return {
        "messages": len(msgs), "matched": sum(1 for m in msgs.values() if m.get("match")),
        "bulk": sum(1 for m in msgs.values() if m.get("bulk")),
        "filed": sum(1 for m in msgs.values() if m.get("outlook_folder")),
        "file_errors": sum(1 for m in msgs.values() if m.get("file_error")),
        "pending": sum(1 for p in props if p["status"] == "pending"),
        "decided": sum(1 for p in props if p["status"] in ("approved", "rejected")),
        "last_pull": state.get("last_pull"), "watermark": state.get("watermark"),
        "settings": get_settings(),
    }


# ---------------------------------------------------------------- auto-poll
def ensure_poller() -> None:
    global _poller_started
    if _poller_started:
        return
    _poller_started = True

    def loop():
        while True:
            minutes = get_settings().get("auto_poll_minutes") or 0
            if minutes <= 0:
                time.sleep(60)
                continue
            try:
                jid = start_pull()
                while (jobs.get(jid) or {}).get("phase") not in ("done", "error"):
                    time.sleep(5)
            except Exception:  # noqa: BLE001
                pass
            time.sleep(minutes * 60)

    threading.Thread(target=loop, daemon=True).start()


# ---------------------------------------------------------------- CLI
if __name__ == "__main__":
    import sys
    cmd = sys.argv[1] if len(sys.argv) > 1 else "report"
    if cmd == "pull":
        days = int(sys.argv[2].rstrip("d")) if len(sys.argv) > 2 and sys.argv[2][0].isdigit() else None
        since = datetime.now() - timedelta(days=days) if days else None
        new = pull(since=since)
        r = categorize(new, with_dates="--no-dates" not in sys.argv)
        _append_mail(new)
        print(f"new: {len(new)} | matched: {r['matched']} | filed: {r['moved']} | proposals: {r['proposals']}")
    elif cmd == "recategorize":
        print(recategorize_all())
    elif cmd == "report":
        for m in sorted(load_mail().values(), key=lambda m: m.get("received") or ""):
            mt = m.get("match") or {}
            doc = m.get("doc") or {}
            print(f"{(m.get('received') or '')[:10]} | {(m.get('sender_name') or '')[:22]:22} | "
                  f"{(m.get('subject') or '')[:50]:50} | {(mt.get('name') or m.get('bulk') or '—')[:34]:34} | "
                  f"{doc.get('label', '')}{' (draft)' if doc.get('draft') else ''}")
    else:
        sys.exit("usage: python -m app.mail pull [7d] [--no-dates] | recategorize | report")
