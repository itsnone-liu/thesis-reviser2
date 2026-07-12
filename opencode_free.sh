#!/usr/bin/env bash
set -euo pipefail

OPENCODE_BIN="${OPENCODE_BIN:-/root/.opencode/bin/opencode}"
MODEL="${OPENCODE_MODEL:-opencode/deepseek-v4-flash-free}"

exec "$OPENCODE_BIN" run -m "$MODEL" "$@"
