#!/bin/bash
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"
export PYTHONUNBUFFERED=1
exec python3 -u -m delivery.batch --workflow civil "$@" >> /tmp/civil_generate.log 2>&1
