"""Engagement advisor: concrete next actions per property, derived from its stage,
Deal Timeline checklist gaps, critical dates, and recent email. Rule-based core is
instant and deterministic; a deeper narrative per property is available via claude -p.
"""

import hashlib
import json
import os
import signal
import subprocess
import threading
import time
from collections import Counter
from datetime import date, datetime

from . import dates, mail, status
from .box_bridge import PROJECT_DIR

ADVICE_PATH = PROJECT_DIR / ".cache" / "status" / "advice.json"
_lock = threading.Lock()

PRIORITY_ORDER = {"critical": 0, "high": 1, "medium": 2, "low": 3}
STALE_PAST_DAYS = 45      # dates further in the past than this are history, not action items
MAX_ACTIONS = 6


def _days_until(iso: str | None) -> int | None:
    if not iso:
        return None
    try:
        return (date.fromisoformat(str(iso)[:10]) - date.today()).days
    except ValueError:
        return None


def _act(actions: list, priority: str, action: str, why: str, due: str | None = None) -> None:
    actions.append({"priority": priority, "action": action, "why": why, "due": due})


def _date_rules(actions: list, row: dict, d: dict) -> None:
    idx = row["stage_idx"]
    closing_days = _days_until(d.get("closing_date"))
    deal_done = idx >= 6 and closing_days is not None and closing_days < 0

    def check(field, label, horizon, priority, text):
        days = _days_until(d.get(field))
        if days is None or deal_done:
            return
        if -STALE_PAST_DAYS <= days < 0:
            _act(actions, "critical", f"Overdue: {text}", f"{label} was {d[field]} ({-days}d ago)", d[field])
        elif 0 <= days <= horizon:
            _act(actions, priority, text, f"{label} {d[field]} — in {days}d", d[field])

    check("deposit_date", "Deposit due", 3, "critical", "Confirm buyer's deposit wired to escrow")
    check("inspection_end", "DD / inspection ends", 7, "critical",
          "DD ends — confirm third-party reports received; deposit goes non-refundable")
    check("deposit_nonrefundable", "Deposit non-refundable", 7, "high", "Deposit goes hard — confirm buyer proceeding")
    check("financing_contingency", "Financing contingency", 7, "high", "Request loan approval docs from buyer's broker")
    if not deal_done and closing_days is not None and -STALE_PAST_DAYS <= closing_days <= 10:
        if closing_days < 0:
            _act(actions, "critical", "Closing date has passed — confirm status or re-paper the date",
                 f"closing was {d['closing_date']} ({-closing_days}d ago)", d.get("closing_date"))
        else:
            _act(actions, "high", "Closing in {}d — request preliminary settlement statement, confirm deed & "
                 "assignment drafted, send notice to tenant".format(closing_days),
                 f"closing {d['closing_date']}", d.get("closing_date"))


def _stage_rules(actions: list, row: dict, d: dict, last_activity_days: int | None) -> None:
    idx, missing = row["stage_idx"], set(row.get("missing") or [])
    closing_days = _days_until(d.get("closing_date"))
    deal_done = idx >= 6 and closing_days is not None and closing_days < 0
    if deal_done:
        if "binder" in missing:
            _act(actions, "low", "Assemble the closing binder (contact sheet, DD docs, third-party & closing docs)",
                 "deal closed but no closing binder filed")
        return
    if idx >= 4:  # PSA / Contract or later
        if "estoppel" in missing:
            _act(actions, "high", "Request tenant estoppel (check the lease for the required window)", "under contract, no estoppel on file")
        if "title" in missing:
            _act(actions, "high", "Request preliminary title report from the escrow agent", "under contract, no title work on file")
        if "deposit" in missing:
            _act(actions, "medium", "Request wiring instructions / confirm the deposit reached escrow", "no deposit or escrow paperwork on file")
        if "reports" in missing:
            _act(actions, "medium", "Confirm buyer ordered third-party reports (Phase 1, survey, PCR, appraisal) — get receipts",
                 "no third-party reports on file")
        if "snda" in missing:
            _act(actions, "low", "If the buyer is financing, request the SNDA", "no SNDA on file")
    elif idx == 3:  # LOI
        if (last_activity_days or 0) >= 14:
            _act(actions, "high", "Chase the seller's first PSA draft", f"LOI stage, no movement for {last_activity_days}d")
        else:
            _act(actions, "medium", "Get the PSA drafted and to the buyer's attorney", "LOI stage — next step is the contract")
    elif idx == 2:  # Prep & Marketing
        if "om" in missing:
            _act(actions, "medium", "Prepare the marketing package / OM", "marketing stage, no OM on file")
        if "lease_abstract" in missing:
            _act(actions, "medium", "Put together the lease abstract", "no lease abstract on file")
        if "financials" in missing:
            _act(actions, "medium", "Verify financials / collect the rent roll", "no financials on file")
    elif idx == 1 and "lease" in missing:
        _act(actions, "medium", "Gather due-diligence documents from the seller into Box (lease, amendments, reports)",
             "listed, but no lease on file")
    if idx >= 1 and "listing_agreement" in missing and idx <= 5:
        _act(actions, "low", "File the signed listing agreement", "not found in the folder")


def _momentum_rules(actions: list, row: dict, last_activity_days: int | None, deal_done: bool) -> None:
    if deal_done or row["stage_idx"] == 0 or last_activity_days is None:
        return
    if last_activity_days >= 30 and not any(a["priority"] in ("critical", "high") for a in actions):
        _act(actions, "medium" if last_activity_days < 60 else "low",
             "Check in — engagement has gone quiet", f"no file or email activity for {last_activity_days}d")


def advise() -> dict:
    st = status.get_status()
    all_dates = dates._load()
    pending = Counter(p["box_path"] for p in mail.list_proposals("pending"))
    mail_last: dict[str, str] = {}
    for m in mail.load_mail().values():
        if m.get("match"):
            bp = m["match"]["box_path"]
            r = m.get("received") or ""
            if r > mail_last.get(bp, ""):
                mail_last[bp] = r

    out = []
    for row in st["properties"]:
        bp = row["box_path"]
        d = all_dates.get(bp, {})
        last = max(filter(None, [row.get("last_activity") or "", mail_last.get(bp, "")]), default="")
        last_days = None
        if last:
            try:
                last_days = (datetime.now() - datetime.fromisoformat(last[:19])).days
            except ValueError:
                pass
        closing_days = _days_until(d.get("closing_date"))
        deal_done = row["stage_idx"] >= 6 and closing_days is not None and closing_days < 0

        actions: list[dict] = []
        if pending.get(bp):
            _act(actions, "high", f"Review {pending[bp]} proposed change(s) from email in the Inbox card",
                 "pending go/no-go decisions")
        _date_rules(actions, row, d)
        _stage_rules(actions, row, d, last_days)
        _momentum_rules(actions, row, last_days, deal_done)

        actions.sort(key=lambda a: (PRIORITY_ORDER[a["priority"]], a.get("due") or "9999"))
        if actions:
            out.append({"box_path": bp, "name": row["name"], "category": row["category"],
                        "stage_idx": row["stage_idx"], "actions": actions[:MAX_ACTIONS],
                        "top_priority": actions[0]["priority"]})
    out.sort(key=lambda p: (PRIORITY_ORDER[p["top_priority"]],
                            (p["actions"][0].get("due") or "9999"), p["name"].lower()))
    return {"properties": out, "generated_at": time.time()}


# ---------------------------------------------------------------- deeper advice (claude -p)
def _inputs_hash(row: dict, d: dict, actions: list, mails: list) -> str:
    blob = json.dumps([row.get("stage_idx"), sorted(row.get("missing") or []),
                       {k: d.get(k) for k, _ in dates.DATE_FIELDS},
                       [a["action"] for a in actions], [m["subject"] for m in mails]], sort_keys=True)
    return hashlib.sha256(blob.encode()).hexdigest()[:16]


def deep_advice(box_path: str) -> dict:
    st = {r["box_path"]: r for r in status.get_status()["properties"]}
    row = st.get(box_path)
    if not row:
        raise KeyError(box_path)
    d = dates._load().get(box_path, {})
    base = next((p for p in advise()["properties"] if p["box_path"] == box_path), None)
    actions = base["actions"] if base else []
    mails = sorted([{"subject": m.get("subject") or "", "received": (m.get("received") or "")[:10],
                     "sender": m.get("sender_name") or "", "excerpt": (m.get("body") or "")[:400]}
                    for m in mail.load_mail().values()
                    if (m.get("match") or {}).get("box_path") == box_path],
                   key=lambda m: m["received"], reverse=True)[:8]

    key = _inputs_hash(row, d, actions, mails)
    cache = {}
    try:
        cache = json.loads(ADVICE_PATH.read_text())
    except (OSError, json.JSONDecodeError):
        pass
    hit = cache.get(box_path)
    if hit and hit.get("hash") == key:
        return hit

    item_label = {i["key"]: i["label"] for i in status.ITEMS}
    prompt = "\n".join([
        "You are a commercial real estate deal advisor. Based ONLY on the state below, give the",
        "broker 3-6 concrete, prioritized next actions for this engagement (imperative voice,",
        "each one line, cite the date or gap it addresses), then one line starting 'Risk:' with",
        "the single biggest risk. No preamble.",
        f"\nPROPERTY: {row['name']} (category: {row['category']})",
        f"STAGE: {status.STAGES[row['stage_idx']-1]['label'] if row['stage_idx'] else 'Not started'}",
        "MISSING CHECKLIST ITEMS: " + (", ".join(item_label.get(k, k) for k in row.get("missing") or []) or "none"),
        "CRITICAL DATES: " + (json.dumps({lbl: d.get(k) for k, lbl in dates.DATE_FIELDS if d.get(k)}) or "{}"),
        ("NOTES: " + d["notes"]) if d.get("notes") else "",
        "RULE-BASED SUGGESTIONS SO FAR: " + ("; ".join(a["action"] for a in actions) or "none"),
        "RECENT EMAIL (newest first):",
        *(f"- [{m['received']}] {m['sender']}: {m['subject']} | {m['excerpt'][:200]}" for m in mails),
        f"\nToday's date: {date.today().isoformat()}",
    ])
    cmd = ["claude", "-p", "--output-format", "json", "--model", os.environ.get("QA_MODEL", "sonnet"),
           "--no-session-persistence", "--tools", "Read", "--permission-mode", "dontAsk"]
    proc = subprocess.Popen(cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                            text=True, start_new_session=True)
    try:
        stdout, stderr = proc.communicate(input=prompt, timeout=120)
    except subprocess.TimeoutExpired:
        os.killpg(os.getpgid(proc.pid), signal.SIGKILL)
        proc.wait()
        raise RuntimeError("claude timed out")
    if proc.returncode != 0:
        raise RuntimeError(f"claude exited {proc.returncode}: {stderr.strip()[:300]}")
    envelope = json.loads(stdout)
    if envelope.get("is_error"):
        raise RuntimeError(str(envelope.get("result"))[:300])

    result = {"hash": key, "text": envelope["result"], "generated_at": time.time()}
    with _lock:
        try:
            cache = json.loads(ADVICE_PATH.read_text())
        except (OSError, json.JSONDecodeError):
            cache = {}
        cache[box_path] = result
        ADVICE_PATH.parent.mkdir(parents=True, exist_ok=True)
        ADVICE_PATH.write_text(json.dumps(cache, indent=1))
    return result
