"""Outreach: AI-drafted, human-approved email to a deal's contacts.

Contacts are derived from the property's matched email history. Claude drafts a
stage-appropriate message (replying inside an existing thread when that fits);
the draft sits in a queue where the user edits and explicitly approves it.
Sending goes through classic Outlook COM (no API) — or into Outlook's Drafts
folder if the user prefers to send it themselves.
"""

import json
import os
import re
import signal
import subprocess
import threading
import time
import uuid
from datetime import date

from . import advisor, dates, mail, status
from .box_bridge import _b64, _ps_args

OUTREACH_PATH = mail.MAIL_DIR / "outreach.json"
_lock = threading.Lock()

DRAFT_TIMEOUT = 120


def _load() -> list[dict]:
    try:
        return json.loads(OUTREACH_PATH.read_text())
    except (OSError, json.JSONDecodeError):
        return []


def _save(items: list[dict]) -> None:
    mail.MAIL_DIR.mkdir(parents=True, exist_ok=True)
    OUTREACH_PATH.write_text(json.dumps(items, indent=1))


# ---------------------------------------------------------------- contacts
def contacts_for(box_path: str) -> list[dict]:
    """People on this deal's email threads: inbox senders, ranked by recency + volume."""
    own = re.compile(r"nathan\.kam@", re.I)
    people: dict[str, dict] = {}
    for m in mail.load_mail().values():
        if (m.get("match") or {}).get("box_path") != box_path:
            continue
        email = (m.get("sender_email") or "").lower()
        if not email or "@" not in email or own.search(email) or mail.BULK_SENDER.search(email):
            continue
        c = people.setdefault(email, {"email": email, "name": m.get("sender_name") or email,
                                      "count": 0, "last": "", "last_subject": "", "entry_id": ""})
        c["count"] += 1
        received = m.get("received") or ""
        if received > c["last"]:
            c["last"], c["last_subject"] = received, m.get("subject") or ""
            c["entry_id"] = m["entry_id"]
    return sorted(people.values(), key=lambda c: (c["last"], c["count"]), reverse=True)


def _recent_threads(box_path: str, n: int = 6) -> list[dict]:
    msgs = [m for m in mail.load_mail().values()
            if (m.get("match") or {}).get("box_path") == box_path]
    msgs.sort(key=lambda m: m.get("received") or "", reverse=True)
    return [{"entry_id": m["entry_id"], "received": (m.get("received") or "")[:10],
             "folder": m.get("folder"), "sender": m.get("sender_name") or "",
             "sender_email": m.get("sender_email") or "", "subject": m.get("subject") or "",
             "excerpt": (m.get("body") or "")[:400]} for m in msgs[:n]]


# ---------------------------------------------------------------- drafting
def draft(box_path: str, action: str = "", instructions: str = "") -> dict:
    rows = {r["box_path"]: r for r in status.get_status()["properties"]}
    row = rows.get(box_path)
    if not row:
        raise KeyError(box_path)
    d = dates._load().get(box_path, {})
    contacts = contacts_for(box_path)
    threads = _recent_threads(box_path)
    base = next((p for p in advisor.advise()["properties"] if p["box_path"] == box_path), None)
    actions = [a["action"] for a in (base["actions"] if base else [])]
    item_label = {i["key"]: i["label"] for i in status.ITEMS}

    prompt = "\n".join(filter(None, [
        "You draft a short, professional email for Nathan Kam, a commercial real estate broker",
        "at Matthews Real Estate Investment Services. Based ONLY on the deal state below, write",
        "the outreach that moves the deal forward. Pick the right recipient from CONTACTS.",
        "If one of the RECENT THREADS is the natural place to continue, set reply_entry_id to",
        "that thread's entry_id (the subject then stays as RE: automatically); otherwise write",
        "a fresh subject. Tone: direct, courteous, concrete — reference specific documents and",
        "dates. 3-8 sentences. Sign off:\n\nBest,\nNathan\n",
        "Return ONLY JSON (no fences):",
        '{"to": "email", "cc": "" , "subject": "…", "body": "…", "reply_entry_id": null,',
        ' "rationale": "one line: why this recipient and this ask"}',
        f"\nPROPERTY: {row['name']} (category {row['category']})",
        f"STAGE: {status.STAGES[row['stage_idx']-1]['label'] if row['stage_idx'] else 'Not started'}",
        "MISSING DOCS: " + (", ".join(item_label.get(k, k) for k in row.get("missing") or []) or "none"),
        "CRITICAL DATES: " + json.dumps({lbl: d.get(k) for k, lbl in dates.DATE_FIELDS if d.get(k)}),
        ("NOTES: " + d["notes"]) if d.get("notes") else "",
        ("PURPOSE OF THIS EMAIL (user-selected action): " + action) if action else
        ("OPEN ACTION ITEMS: " + ("; ".join(actions) or "general check-in")),
        ("EXTRA INSTRUCTIONS FROM NATHAN: " + instructions) if instructions else "",
        "\nCONTACTS:",
        *(f"- {c['name']} <{c['email']}> — {c['count']} emails, last {c['last'][:10]}: {c['last_subject'][:70]}"
          for c in contacts[:8]),
        "(none on file — leave \"to\" empty and say so in rationale)" if not contacts else "",
        "\nRECENT THREADS (newest first):",
        *(f"- entry_id={t['entry_id'][-24:]} [{t['received']}] {t['sender']}: {t['subject'][:70]} | {t['excerpt'][:150]}"
          for t in threads),
        f"\nToday: {date.today().isoformat()}",
    ]))

    cmd = ["claude", "-p", "--output-format", "json", "--model", os.environ.get("QA_MODEL", "sonnet"),
           "--no-session-persistence", "--tools", "Read", "--permission-mode", "dontAsk"]
    proc = subprocess.Popen(cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                            text=True, start_new_session=True)
    try:
        stdout, stderr = proc.communicate(input=prompt, timeout=DRAFT_TIMEOUT)
    except subprocess.TimeoutExpired:
        os.killpg(os.getpgid(proc.pid), signal.SIGKILL)
        proc.wait()
        raise RuntimeError("claude timed out drafting")
    if proc.returncode != 0:
        raise RuntimeError(f"claude exited {proc.returncode}: {stderr.strip()[:300]}")
    envelope = json.loads(stdout)
    if envelope.get("is_error"):
        raise RuntimeError(str(envelope.get("result"))[:300])
    match = re.search(r"\{.*\}", envelope["result"], re.S)
    if not match:
        raise RuntimeError(f"no JSON in draft output: {envelope['result'][:200]}")
    parsed = json.loads(match.group(0))

    # short entry-id suffixes back to full ids
    rid = parsed.get("reply_entry_id")
    if rid:
        full = next((t["entry_id"] for t in threads if t["entry_id"].endswith(str(rid))), None)
        parsed["reply_entry_id"] = full

    item = {
        "id": uuid.uuid4().hex[:10], "box_path": box_path, "property_name": row["name"],
        "action": action or None, "status": "pending", "created": time.time(),
        "to": parsed.get("to") or "", "cc": parsed.get("cc") or "",
        "subject": parsed.get("subject") or f"{row['name']}", "body": parsed.get("body") or "",
        "reply_entry_id": parsed.get("reply_entry_id"),
        "reply_subject": next((t["subject"] for t in threads
                               if t["entry_id"] == parsed.get("reply_entry_id")), None),
        "rationale": parsed.get("rationale") or "",
        "contacts": contacts[:8],
    }
    with _lock:
        items = _load()
        items.append(item)
        _save(items)
    return item


# ---------------------------------------------------------------- approve / send
def _dispatch(item: dict, mode: str) -> dict:
    spec = {"mode": mode, "to": item["to"], "cc": item["cc"], "subject": item["subject"],
            "body": item["body"], "reply_entry_id": item.get("reply_entry_id")}
    proc = subprocess.run(_ps_args("outlook_send.ps1", "-SpecB64", _b64(json.dumps(spec))),
                          capture_output=True, text=True, timeout=120)
    out = proc.stdout.strip().lstrip("﻿")
    if not out:
        return {"ok": False, "error": (proc.stderr.strip() or "no output from Outlook")[:300]}
    try:
        return json.loads(out.splitlines()[-1])
    except json.JSONDecodeError:
        return {"ok": False, "error": out[:300]}


def decide(outreach_id: str, decision: str, edits: dict | None = None) -> dict:
    """decision: 'send' | 'outlook_draft' | 'discard'. Edits (to/cc/subject/body) apply first."""
    with _lock:
        items = _load()
        item = next((i for i in items if i["id"] == outreach_id), None)
        if not item:
            raise KeyError(outreach_id)
        if item["status"] != "pending":
            return item
        for k in ("to", "cc", "subject", "body"):
            if edits and edits.get(k) is not None:
                item[k] = edits[k]
        if decision == "discard":
            item["status"] = "discarded"
        else:
            if not item["to"].strip() and decision == "send":
                raise ValueError("no recipient — fill in To before sending")
            result = _dispatch(item, "send" if decision == "send" else "draft")
            item["result"] = result
            if result.get("ok"):
                item["status"] = "sent" if decision == "send" else "outlook_draft"
            else:
                item["status"] = "pending"  # keep editable; surface the error
        item["decided_at"] = time.time()
        _save(items)
        return item


def list_outreach() -> dict:
    items = _load()
    return {"pending": [i for i in items if i["status"] == "pending"],
            "history": sorted([i for i in items if i["status"] != "pending"],
                              key=lambda i: -(i.get("decided_at") or 0))[:40]}
