#!/usr/bin/env python3
"""One-off: group Box root property folders into the 0-* category folders.

Dry-run by default; pass --apply to actually move. Names are resolved against
the live Box root via the UTF-8 JSON bridge (console output mangles non-ASCII),
matched by punctuation-insensitive normalization. Unmatched names are skipped
and reported — never guessed.
"""

import argparse
import json
import re
import subprocess
import sys
import unicodedata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import box_bridge
from app.box_bridge import _b64, _ps_args, get_box_root, list_children  # noqa: E402

from app.categories import MAPPING, squash  # noqa: E402



def resolve_moves(root: str) -> tuple[list[tuple[str, str]], list[str]]:
    real_dirs = [i["name"] for i in list_children(root) if i["is_dir"]]
    by_key: dict[str, list[str]] = {}
    for n in real_dirs:
        by_key.setdefault(squash(n), []).append(n)

    moves, problems = [], []
    for category, names in MAPPING.items():
        if category not in real_dirs:
            problems.append(f"category folder missing in Box root: {category}")
            continue
        for want in names:
            candidates = by_key.get(squash(want), [])
            if len(candidates) == 1:
                moves.append((candidates[0], category))
            elif not candidates:
                problems.append(f"no match in Box root (skipped): {want}")
            else:
                problems.append(f"ambiguous match (skipped): {want} -> {candidates}")
    return moves, problems


def move_one(root: str, name: str, category: str) -> dict:
    src = f"{root}\\{name}"
    dst = f"{root}\\{category}\\{name}"
    proc = subprocess.run(
        _ps_args("move_folder.ps1", "-SrcB64", _b64(src), "-DstB64", _b64(dst)),
        capture_output=True, text=True, timeout=180,
    )
    out = proc.stdout.strip().lstrip("﻿")
    if not out:
        return {"ok": False, "error": (proc.stderr.strip() or "no output")[:300]}
    return json.loads(out)


def update_settings(root: str, moved: dict[str, str]) -> int:
    """Rewrite saved index paths whose folder (or ancestor) moved."""
    settings_path = box_bridge.PROJECT_DIR / ".cache" / "settings.json"
    try:
        settings = json.loads(settings_path.read_text())
    except (OSError, json.JSONDecodeError):
        return 0
    changed = 0
    new_paths = []
    for p in settings.get("paths", []):
        for old, new in moved.items():
            if p == old or p.lower().startswith(old.lower() + "\\"):
                p = new + p[len(old):]
                changed += 1
                break
        new_paths.append(p)
    if changed:
        settings["paths"] = new_paths
        settings_path.write_text(json.dumps(settings, indent=2))
    return changed


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--apply", action="store_true", help="actually move (default: dry run)")
    args = ap.parse_args()

    root = get_box_root()
    moves, problems = resolve_moves(root)

    counts: dict[str, int] = {}
    for name, category in moves:
        counts[category] = counts.get(category, 0) + 1
        print(f"  {name}  ->  {category}")
    print(f"\nPlanned moves: {len(moves)}  " +
          " ".join(f"[{c}: {n}]" for c, n in sorted(counts.items())))
    if problems:
        print(f"\nSkipped ({len(problems)}):")
        for p in problems:
            print(f"  ! {p}")

    if not args.apply:
        print("\nDry run only — re-run with --apply to move.")
        return

    print("\nApplying…")
    failures, moved = [], {}
    for i, (name, category) in enumerate(moves, 1):
        result = move_one(root, name, category)
        if result.get("ok"):
            moved[f"{root}\\{name}"] = f"{root}\\{category}\\{name}"
            print(f"  [{i}/{len(moves)}] OK    {name} -> {category}")
        else:
            failures.append((name, result.get("error", "?")))
            print(f"  [{i}/{len(moves)}] FAIL  {name}: {result.get('error', '?')}")

    n_settings = update_settings(root, moved)
    print(f"\nDone: {len(moved)} moved, {len(failures)} failed, "
          f"{n_settings} saved index path(s) updated.")
    if failures:
        print("Failed folders (left in place):")
        for name, err in failures:
            print(f"  ! {name}: {err}")


if __name__ == "__main__":
    main()
