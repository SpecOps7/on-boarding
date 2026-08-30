#!/usr/bin/env bash
# Start the Box Q&A app. Run inside WSL: bash run.sh [port]
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="${VENV_DIR:-$HOME/.venvs/onboarding}"
PORT="${1:-8712}"

[ -x "$VENV_DIR/bin/uvicorn" ] || { echo "Venv not ready — run: bash setup.sh" >&2; exit 1; }

echo "Box Q&A running at http://localhost:$PORT  (Ctrl+C to stop)"
# Open the default Windows browser (best effort)
(sleep 1 && powershell.exe -NoProfile -Command "Start-Process 'http://localhost:$PORT'" >/dev/null 2>&1) &

cd "$PROJECT_DIR"
exec "$VENV_DIR/bin/uvicorn" app.main:app --host 127.0.0.1 --port "$PORT"
