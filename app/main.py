"""Box Q&A web app: FastAPI routes + static UI."""

import asyncio
import json
import threading

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel

from . import advisor, box_bridge, dates, indexer, jobs, mail, outreach, qa, status
from .box_bridge import PROJECT_DIR

app = FastAPI(title="Box Q&A")

_ask_lock = asyncio.Lock()  # serialize claude -p calls
_settings_lock = threading.Lock()

SETTINGS_PATH = PROJECT_DIR / ".cache" / "settings.json"


def load_settings() -> dict:
    try:
        return json.loads(SETTINGS_PATH.read_text())
    except (OSError, json.JSONDecodeError):
        return {"paths": []}


def save_settings(settings: dict) -> None:
    SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
    SETTINGS_PATH.write_text(json.dumps(settings, indent=2))


def remember_path(box_path: str) -> None:
    with _settings_lock:
        settings = load_settings()
        if box_path not in settings["paths"]:
            settings["paths"].append(box_path)
            save_settings(settings)


class IndexRequest(BaseModel):
    box_path: str


class AskRequest(BaseModel):
    slug: str
    question: str


class SettingsRequest(BaseModel):
    paths: list[str]


class OverrideRequest(BaseModel):
    box_path: str
    stage: str | None = None


class DecisionRequest(BaseModel):
    accept: list[int] | None = None


class AssignRequest(BaseModel):
    box_path: str


class MailSettingsRequest(BaseModel):
    auto_poll_minutes: int | None = None
    since_days: int | None = None


@app.on_event("startup")
def _start_mail_poller():
    mail.ensure_poller()


@app.get("/")
def home():
    return FileResponse(PROJECT_DIR / "static" / "index.html")


@app.get("/api/box/roots")
def box_roots():
    try:
        return {"root": box_bridge.get_box_root(), "items": box_bridge.list_box_roots()}
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f"Box listing failed: {e}")


@app.get("/api/box/children")
def box_children(path: str):
    root = box_bridge.get_box_root()
    if not path.lower().startswith(root.lower()):
        raise HTTPException(400, "path must be inside the Box root")
    try:
        return {"items": box_bridge.list_children(path)}
    except Exception as e:  # noqa: BLE001
        raise HTTPException(502, f"Box listing failed: {e}")


def _require_in_box(box_path: str) -> None:
    root = box_bridge.get_box_root()
    if not box_path.lower().startswith(root.lower()):
        raise HTTPException(400, "path must be inside the Box root")


def _index_one(job_id: str, box_path: str, note: str = "") -> str:
    """Sync + extract one folder, streaming progress into the job. Returns a summary line."""
    slug = indexer.make_slug(box_path)
    jobs.update(job_id, phase="sync", note=note, done=0, total=0, current_file="")

    def sync_progress(event):
        if event.get("event") == "start":
            jobs.update(job_id, total=event.get("total", 0), done=0)
        elif event.get("event") == "file":
            job = jobs.get(job_id)
            jobs.update(job_id, done=(job["done"] + 1) if job else 0,
                        current_file=event.get("file", ""))

    summary = box_bridge.sync_folder(box_path, slug, progress_cb=sync_progress)
    jobs.update(job_id, phase="extract", done=0, total=0, current_file="")
    result = indexer.index_folder(slug, box_path, progress_cb=lambda p: jobs.update(job_id, **p))
    return (f"synced {summary['copied']} new / {summary['skipped']} unchanged "
            f"({summary['excluded']} excluded), {result['chunks']} chunks "
            f"from {result['files']} files")


@app.post("/api/index")
def start_index(req: IndexRequest):
    _require_in_box(req.box_path)
    slug = indexer.make_slug(req.box_path)
    remember_path(req.box_path)
    job_id = jobs.create("index", slug=slug, box_path=req.box_path)

    def work():
        line = _index_one(job_id, req.box_path)
        jobs.update(job_id, message=line.capitalize() + ".")

    jobs.run_in_thread(job_id, work)
    return {"job_id": job_id, "slug": slug}


@app.get("/api/settings")
def get_settings():
    return load_settings()


@app.post("/api/settings")
def set_settings(req: SettingsRequest):
    paths = []
    for p in dict.fromkeys(req.paths):  # dedupe, keep order
        _require_in_box(p)
        paths.append(p)
    with _settings_lock:
        settings = load_settings()
        settings["paths"] = paths
        save_settings(settings)
    return {"paths": paths}


@app.post("/api/index-all")
def index_all():
    paths = load_settings()["paths"]
    if not paths:
        raise HTTPException(400, "no folders selected — pick some in Settings first")
    job_id = jobs.create("index-all")

    def work():
        lines = []
        for i, box_path in enumerate(paths, 1):
            name = box_path.rstrip("\\").split("\\")[-1]
            line = _index_one(job_id, box_path, note=f"[{i}/{len(paths)}] {name}")
            lines.append(f"{name}: {line}")
        jobs.update(job_id, note="", message=" · ".join(lines))

    jobs.run_in_thread(job_id, work)
    return {"job_id": job_id}


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str):
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(404, "no such job")
    return job


@app.get("/api/folders")
def folders():
    return {"folders": indexer.list_folders()}


@app.get("/api/folders/{slug}/files")
def folder_files(slug: str):
    info = indexer.folder_info(slug)
    if not info:
        raise HTTPException(404, "folder not indexed")
    return {"info": info, "files": indexer.folder_files(slug)}


@app.post("/api/ask")
async def ask(req: AskRequest):
    if indexer.folder_info(req.slug) is None:
        raise HTTPException(404, "folder not indexed")
    if not req.question.strip():
        raise HTTPException(400, "empty question")
    async with _ask_lock:
        try:
            return await asyncio.to_thread(qa.answer, req.slug, req.question.strip())
        except Exception as e:  # noqa: BLE001
            raise HTTPException(502, f"answer failed: {e}")


@app.get("/dashboard")
def dashboard():
    return FileResponse(PROJECT_DIR / "static" / "dashboard.html")


@app.get("/api/status")
def pipeline_status():
    return status.get_status()


@app.post("/api/status/scan")
def start_status_scan():
    return {"job_id": status.start_scan()}


@app.post("/api/status/override")
def status_override(req: OverrideRequest):
    try:
        status.set_override(req.box_path, req.stage)
    except ValueError as e:
        raise HTTPException(400, str(e))
    return {"ok": True}


@app.get("/api/dates")
def critical_dates():
    return dates.get_dates()


@app.post("/api/dates/extract")
def start_dates_extract(force: bool = False):
    return {"job_id": dates.start_extract(force=force)}


@app.get("/api/mail/summary")
def mail_summary():
    return mail.summary()


@app.post("/api/mail/pull")
def mail_pull():
    return {"job_id": mail.start_pull()}


@app.post("/api/mail/recategorize")
def mail_recategorize():
    return {"job_id": mail.start_recategorize()}


@app.get("/api/mail/proposals")
def mail_proposals(status_filter: str = "pending"):
    return {"proposals": mail.list_proposals(status_filter),
            "stages": [s["label"] for s in status.STAGES]}


@app.post("/api/mail/proposals/{pid}/approve")
def mail_approve(pid: str, req: DecisionRequest):
    try:
        return mail.decide(pid, "approve", req.accept)
    except KeyError:
        raise HTTPException(404, "no such proposal")


@app.post("/api/mail/proposals/{pid}/reject")
def mail_reject(pid: str):
    try:
        return mail.decide(pid, "reject")
    except KeyError:
        raise HTTPException(404, "no such proposal")


@app.get("/api/mail/unmatched")
def mail_unmatched():
    return {"unmatched": mail.list_unmatched()}


@app.get("/api/mail/filed")
def mail_filed():
    return {"filed": mail.list_filed()}


@app.post("/api/mail/assign/{entry_id}")
def mail_assign(entry_id: str, req: AssignRequest):
    _require_in_box(req.box_path)
    try:
        prop = mail.assign(entry_id, req.box_path)
    except KeyError as e:
        raise HTTPException(404, f"unknown email or property: {e}")
    return {"proposal": prop}


@app.get("/api/advice")
def advice():
    return advisor.advise()


class DeepAdviceRequest(BaseModel):
    box_path: str


@app.post("/api/advice/deep")
async def advice_deep(req: DeepAdviceRequest):
    async with _ask_lock:
        try:
            return await asyncio.to_thread(advisor.deep_advice, req.box_path)
        except KeyError:
            raise HTTPException(404, "unknown property")
        except Exception as e:  # noqa: BLE001
            raise HTTPException(502, f"advice failed: {e}")


class DraftRequest(BaseModel):
    box_path: str
    action: str = ""
    instructions: str = ""


class OutreachDecision(BaseModel):
    to: str | None = None
    cc: str | None = None
    subject: str | None = None
    body: str | None = None


@app.get("/api/outreach")
def outreach_list():
    return outreach.list_outreach()


@app.post("/api/outreach/draft")
async def outreach_draft(req: DraftRequest):
    _require_in_box(req.box_path)
    async with _ask_lock:
        try:
            return await asyncio.to_thread(outreach.draft, req.box_path, req.action, req.instructions)
        except KeyError:
            raise HTTPException(404, "unknown property")
        except Exception as e:  # noqa: BLE001
            raise HTTPException(502, f"drafting failed: {e}")


@app.post("/api/outreach/{oid}/{decision}")
def outreach_decide(oid: str, decision: str, req: OutreachDecision):
    if decision not in ("send", "outlook_draft", "discard"):
        raise HTTPException(400, "decision must be send | outlook_draft | discard")
    try:
        return outreach.decide(oid, decision, req.model_dump())
    except KeyError:
        raise HTTPException(404, "no such draft")
    except ValueError as e:
        raise HTTPException(400, str(e))


@app.get("/api/outreach/contacts")
def outreach_contacts(box_path: str):
    _require_in_box(box_path)
    return {"contacts": outreach.contacts_for(box_path)}


@app.get("/api/mail/settings")
def mail_settings():
    return mail.get_settings()


@app.post("/api/mail/settings")
def mail_settings_set(req: MailSettingsRequest):
    return mail.set_settings({k: v for k, v in req.model_dump().items() if v is not None})


@app.post("/api/folders/{slug}/clear-history")
def clear_history(slug: str):
    qa.clear_history(slug)
    return {"ok": True}
