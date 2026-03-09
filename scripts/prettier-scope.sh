#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

paths=(
  "package.json"
  "tsconfig.json"
  "tsconfig.renderer.json"
  "tsconfig.typecheck.json"
  ".prettierrc.json"
  ".github"
  "build"
  "launcher"
  "scripts"
  "src"
)

BUN_BIN="$(command -v bun.exe || command -v bun)"
exec "$BUN_BIN" x prettier "$@" "${paths[@]}"
