#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

paths=(
  "package.json"
  "tsconfig.json"
  "tsconfig.renderer.json"
  "tsconfig.typecheck.json"
  ".github"
  "packaging/bun/bun-runtime-manifest.json"
  "packaging/bun/bun-source-manifest.json"
  "e2e"
  "launcher"
  "scripts"
  "src"
)

BUN_BIN="$(command -v bun.exe || command -v bun)"
exec "$BUN_BIN" x prettier "$@" "${paths[@]}"
