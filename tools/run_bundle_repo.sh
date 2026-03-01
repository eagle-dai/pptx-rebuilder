#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if command -v uv >/dev/null 2>&1; then
  uv run "${SCRIPT_DIR}/bundle_repo.py" "$@"
else
  python3 "${SCRIPT_DIR}/bundle_repo.py" "$@"
fi
