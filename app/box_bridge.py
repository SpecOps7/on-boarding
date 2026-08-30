"""PowerShell interop with the Windows-mounted Box Drive.

WSL cannot traverse the Box Drive reparse point directly, so every touch of
Box goes through powershell.exe helper scripts. Paths are passed base64-encoded
so apostrophes/ampersands/brackets in Box folder names can never break parsing.
"""

import base64
import json
import os
import subprocess
import threading
import time
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent.parent
PS_DIR = PROJECT_DIR / "app" / "ps"
CACHE_BOX_DIR = PROJECT_DIR / ".cache" / "box"

LIST_TIMEOUT = 60
SYNC_STALL_TIMEOUT = 300  # kill sync if no progress line for this long
SYNC_TOTAL_TIMEOUT = 30 * 60

_box_root: str | None = None


def wsl_to_win(path: str | Path) -> str:
    """Convert a WSL path to a Windows path (handles /mnt/* and \\\\wsl$ paths)."""
    proc = subprocess.run(
        ["wslpath", "-w", str(Path(path).resolve())],
        capture_output=True, text=True, timeout=10,
    )
    if proc.returncode != 0:
        raise ValueError(f"wslpath failed for {path}: {proc.stderr.strip()}")
    return proc.stdout.strip()


def get_box_root() -> str:
    """Box Drive root: $BOX_ROOT if set, else <Windows user profile>\\Box."""
    global _box_root
    if _box_root is None:
        env = os.environ.get("BOX_ROOT")
        if env:
            _box_root = env
        else:
            proc = subprocess.run(
                ["powershell.exe", "-NoProfile", "-Command", "Write-Output $env:USERPROFILE"],
                capture_output=True, text=True, timeout=30,
            )
            profile = proc.stdout.strip()
            if proc.returncode != 0 or not profile:
                raise RuntimeError("could not resolve Windows user profile; set BOX_ROOT")
            _box_root = profile + r"\Box"
    return _box_root


def _b64(s: str) -> str:
    return base64.b64encode(s.encode("utf-8")).decode("ascii")


def _ps_args(script: str, *args: str) -> list[str]:
    return [
        "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", wsl_to_win(PS_DIR / script), *args,
    ]


def _parse_json(stdout: str):
    return json.loads(stdout.strip().lstrip("\ufeff"))


def list_children(box_path: str) -> list[dict]:
    """Non-recursive listing of one Box directory: [{name, is_dir, size, mtime_iso, ext}]."""
    proc = subprocess.run(
        _ps_args("list_children.ps1", "-PathB64", _b64(box_path)),
        capture_output=True, text=True, timeout=LIST_TIMEOUT,
    )
    if proc.returncode != 0:
        raise RuntimeError(f"list_children failed for {box_path!r}: {proc.stderr.strip()[:500]}")
    items = _parse_json(proc.stdout)
    return sorted(items, key=lambda i: (not i["is_dir"], i["name"].lower()))


def list_box_roots() -> list[dict]:
    return list_children(get_box_root())


def sync_folder(box_path: str, slug: str, progress_cb=None) -> dict:
    """Copy a Box folder into .cache/box/<slug>/ incrementally.

    Streams per-file progress events to progress_cb(event_dict). Returns the
    final summary {total, copied, skipped, excluded, errors}.
    """
    dest = CACHE_BOX_DIR / slug
    dest.mkdir(parents=True, exist_ok=True)

    proc = subprocess.Popen(
        _ps_args("sync_folder.ps1", "-SrcB64", _b64(box_path), "-DstB64", _b64(wsl_to_win(dest))),
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
        start_new_session=True,
    )

    # Watchdog: Box hydration can stall on a bad network; kill if no output for a while.
    last_activity = [time.monotonic()]
    start = last_activity[0]
    stop = threading.Event()

    def watchdog():
        while not stop.wait(10):
            now = time.monotonic()
            if now - last_activity[0] > SYNC_STALL_TIMEOUT or now - start > SYNC_TOTAL_TIMEOUT:
                proc.kill()
                return

    threading.Thread(target=watchdog, daemon=True).start()

    summary = None
    try:
        for line in proc.stdout:
            last_activity[0] = time.monotonic()
            line = line.strip().lstrip("\ufeff")
            if not line:
                continue
            try:
                event = json.loads(line)
            except json.JSONDecodeError:
                continue
            if event.get("event") == "done":
                summary = event
            if progress_cb:
                progress_cb(event)
        proc.wait(timeout=30)
    finally:
        stop.set()

    if summary is None:
        stderr = proc.stderr.read()[:500] if proc.stderr else ""
        raise RuntimeError(f"sync of {box_path!r} did not complete (stalled or killed). {stderr}")
    return summary
