#!/usr/bin/env bash
set -euo pipefail

OPENCODE_BIN="${OPENCODE_BIN:-/root/.opencode/bin/opencode}"
MODEL="${OPENCODE_MODEL:-opencode/deepseek-v4-flash-free}"

if ! command -v "$OPENCODE_BIN" >/dev/null 2>&1; then
  echo "opencode binary not found: $OPENCODE_BIN" >&2
  exit 1
fi

exec "$OPENCODE_BIN" run -m "$MODEL" "$@"
