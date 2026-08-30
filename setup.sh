#!/usr/bin/env bash
# One-time setup for the Box Q&A app. Run inside WSL: bash setup.sh
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="${VENV_DIR:-$HOME/.venvs/onboarding}"

say()  { printf '\033[1;34m==>\033[0m %s\n' "$*"; }
fail() { printf '\033[1;31mERROR:\033[0m %s\n' "$*" >&2; exit 1; }

# --- 0. Must be WSL with Windows interop (the app reads Box Drive through powershell.exe)
grep -qi microsoft /proc/version 2>/dev/null || fail "This app must run inside WSL on the machine where Box Drive is installed."
command -v powershell.exe >/dev/null 2>&1 || fail "powershell.exe not reachable from WSL (Windows interop is disabled?)."

# --- 1. Python 3.10+ (install via apt if missing)
PY=python3
if ! command -v python3 >/dev/null 2>&1 || ! python3 -c 'import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)'; then
  say "Python 3.10+ not found — installing via apt (sudo will prompt)…"
  sudo apt-get update -y
  sudo apt-get install -y python3 python3-venv python3-pip
fi
python3 -c 'import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)' || fail "python3 is older than 3.10 even after install."
# venv module is a separate package on Debian/Ubuntu
python3 -m venv --help >/dev/null 2>&1 || { say "Installing python3-venv…"; sudo apt-get install -y python3-venv; }

# --- 2. Virtualenv on WSL's native filesystem (venvs on /mnt/c are very slow)
if [ ! -x "$VENV_DIR/bin/python" ]; then
  say "Creating venv at $VENV_DIR"
  python3 -m venv "$VENV_DIR"
fi
say "Installing Python dependencies"
"$VENV_DIR/bin/pip" install --quiet --upgrade pip
"$VENV_DIR/bin/pip" install --quiet -r "$PROJECT_DIR/requirements.txt"

# --- 3. Claude Code CLI (powers the answers; no API key needed)
if ! command -v claude >/dev/null 2>&1; then
  say "Claude Code CLI not found — installing…"
  curl -fsSL https://claude.ai/install.sh | bash
  command -v claude >/dev/null 2>&1 || fail "claude still not on PATH — open a new shell or install manually, then re-run setup."
fi
say "claude CLI: $(claude --version 2>/dev/null | head -1)"

# --- 4. Box Drive reachable?
BOX_ROOT_DETECTED="$(powershell.exe -NoProfile -Command 'Write-Output $env:USERPROFILE' | tr -d '\r')\\Box"
BOX_ROOT="${BOX_ROOT:-$BOX_ROOT_DETECTED}"
if powershell.exe -NoProfile -Command "Test-Path -LiteralPath '$BOX_ROOT'" | grep -qi true; then
  say "Box Drive found at $BOX_ROOT"
else
  printf '\033[1;33mWARNING:\033[0m Box Drive not found at %s.\n' "$BOX_ROOT"
  echo "  Install Box Drive on Windows (https://www.box.com/resources/downloads),"
  echo "  or set BOX_ROOT to its location before starting the app:  export BOX_ROOT='D:\\SomePath\\Box'"
fi

say "Setup complete. Start the app with:  bash run.sh"
