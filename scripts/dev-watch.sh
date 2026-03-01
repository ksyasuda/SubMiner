#!/usr/bin/env bash

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

cd "$ROOT_DIR"

electron_args=("$@")
if [[ ${#electron_args[@]} -eq 0 ]]; then
  electron_args=(--start --dev)
fi

if ! command -v bun >/dev/null 2>&1; then
  echo "[ERROR] bun not found in PATH" >&2
  exit 1
fi

TS_WATCH_PID=""
RENDER_WATCH_PID=""

cleanup() {
  local pids=("$TS_WATCH_PID" "$RENDER_WATCH_PID")
  for pid in "${pids[@]}"; do
    if [[ -n "$pid" ]] && kill -0 "$pid" 2>/dev/null; then
      kill "$pid" 2>/dev/null || true
    fi
  done
}

trap cleanup EXIT INT TERM

sync_renderer_assets() {
  mkdir -p dist/renderer
  cp src/renderer/index.html src/renderer/style.css dist/renderer/
  mkdir -p dist/renderer/fonts
  cp -R src/renderer/fonts/. dist/renderer/fonts/
}

echo "[INFO] Syncing renderer static assets"
sync_renderer_assets

echo "[INFO] Running initial compile"
bun run tsc
bun run build:renderer

echo "[INFO] Starting TypeScript watch"
bun run tsc --watch --preserveWatchOutput &
TS_WATCH_PID=$!

echo "[INFO] Starting renderer watch"
bunx esbuild src/renderer/renderer.ts \
  --bundle \
  --platform=browser \
  --format=esm \
  --target=es2022 \
  --outfile=dist/renderer/renderer.js \
  --sourcemap \
  --watch &
RENDER_WATCH_PID=$!

echo "[INFO] Launching Electron with args: ${electron_args[*]}"
bun run electron . "${electron_args[@]}"
