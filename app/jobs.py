"""In-memory background job registry with progress, for index runs."""

import threading
import time
import uuid

_jobs: dict[str, dict] = {}
_lock = threading.Lock()


def create(kind: str, **fields) -> str:
    job_id = uuid.uuid4().hex[:12]
    with _lock:
        _jobs[job_id] = {
            "id": job_id, "kind": kind, "phase": "queued",
            "done": 0, "total": 0, "current_file": "", "message": "",
            "error": None, "started": time.time(), **fields,
        }
    return job_id


def update(job_id: str, **fields) -> None:
    with _lock:
        if job_id in _jobs:
            _jobs[job_id].update(fields)


def get(job_id: str) -> dict | None:
    with _lock:
        job = _jobs.get(job_id)
        return dict(job) if job else None


def run_in_thread(job_id: str, target) -> None:
    def wrapper():
        try:
            target()
            update(job_id, phase="done")
        except Exception as e:  # noqa: BLE001 - surface to UI instead of dying silently
            update(job_id, phase="error", error=f"{type(e).__name__}: {e}"[:500])

    threading.Thread(target=wrapper, daemon=True).start()
