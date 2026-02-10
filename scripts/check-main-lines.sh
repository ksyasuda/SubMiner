#!/usr/bin/env bash
set -euo pipefail

target="${1:-1500}"
file="${2:-src/main.ts}"

if [[ ! -f "$file" ]]; then
  echo "[ERROR] File not found: $file" >&2
  exit 1
fi

if ! [[ "$target" =~ ^[0-9]+$ ]]; then
  echo "[ERROR] Target line count must be an integer. Got: $target" >&2
  exit 1
fi

actual="$(wc -l < "$file" | tr -d ' ')"

echo "[INFO] $file lines: $actual (target: <= $target)"
if (( actual > target )); then
  echo "[ERROR] Line gate failed: $actual > $target" >&2
  exit 1
fi

echo "[OK] Line gate passed"
