#!/usr/bin/env bash
set -euo pipefail

usage() {
  cat <<'EOF'
Usage: classify_subminer_diff.sh [path ...]

Emit suggested verification lanes for explicit paths or current local git changes.

Output format:
  lane:<name>
  flag:<name>
  reason:<text>
EOF
}

has_item() {
  local needle=$1
  shift || true
  local item
  for item in "$@"; do
    if [[ "$item" == "$needle" ]]; then
      return 0
    fi
  done
  return 1
}

add_lane() {
  local lane=$1
  if ! has_item "$lane" "${LANES[@]:-}"; then
    LANES+=("$lane")
  fi
}

add_flag() {
  local flag=$1
  if ! has_item "$flag" "${FLAGS[@]:-}"; then
    FLAGS+=("$flag")
  fi
}

add_reason() {
  REASONS+=("$1")
}

collect_git_paths() {
  local top_level
  if ! top_level=$(git rev-parse --show-toplevel 2>/dev/null); then
    return 0
  fi

  (
    cd "$top_level"
    if git rev-parse --verify HEAD >/dev/null 2>&1; then
      git diff --name-only --relative HEAD --
      git diff --name-only --relative --cached --
    else
      git diff --name-only --relative --
      git diff --name-only --relative --cached --
    fi
    git ls-files --others --exclude-standard
  ) | awk 'NF' | sort -u
}

if [[ "${1:-}" == "--help" || "${1:-}" == "-h" ]]; then
  usage
  exit 0
fi

declare -a PATHS=()
declare -a LANES=()
declare -a FLAGS=()
declare -a REASONS=()

if [[ $# -gt 0 ]]; then
  while [[ $# -gt 0 ]]; do
    PATHS+=("$1")
    shift
  done
else
  while IFS= read -r line; do
    [[ -n "$line" ]] && PATHS+=("$line")
  done < <(collect_git_paths)
fi

if [[ ${#PATHS[@]} -eq 0 ]]; then
  add_lane "core"
  add_reason "no changed paths detected -> default to core"
fi

for path in "${PATHS[@]}"; do
  specialized=0

  case "$path" in
    docs-site/*|docs/*|changes/*|README.md)
      add_lane "docs"
      add_reason "$path -> docs"
      specialized=1
      ;;
  esac

  case "$path" in
    src/config/*|src/generate-config-example.ts|src/verify-config-example.ts|docs-site/public/config.example.jsonc|config.example.jsonc)
      add_lane "config"
      add_reason "$path -> config"
      specialized=1
      ;;
  esac

  case "$path" in
    stats/*)
      add_lane "stats"
      add_reason "$path -> stats"
      specialized=1
      ;;
  esac

  case "$path" in
    launcher/*|plugin/subminer/*|plugin/subminer.conf|scripts/test-plugin-*|scripts/get-mpv-window-*|scripts/configure-plugin-binary-path.mjs)
      add_lane "launcher-plugin"
      add_reason "$path -> launcher-plugin"
      add_flag "real-runtime-candidate"
      add_reason "$path -> real-runtime-candidate"
      specialized=1
      ;;
  esac

  case "$path" in
    src/main.ts|src/main-entry.ts|src/preload.ts|src/main/*|src/core/services/mpv*|src/core/services/overlay*|src/renderer/*|src/window-trackers/*|scripts/prepare-build-assets.mjs)
      add_lane "runtime-compat"
      add_reason "$path -> runtime-compat"
      add_flag "real-runtime-candidate"
      add_reason "$path -> real-runtime-candidate"
      specialized=1
      ;;
  esac

  if [[ "$specialized" == "0" ]]; then
    case "$path" in
      src/*|package.json|tsconfig*.json|scripts/*|Makefile)
        add_lane "core"
        add_reason "$path -> core"
        ;;
    esac
  fi

  case "$path" in
    package.json|src/main.ts|src/main-entry.ts|src/preload.ts)
      add_flag "broad-impact"
      add_reason "$path -> broad-impact"
      ;;
  esac
done

if [[ ${#LANES[@]} -eq 0 ]]; then
  add_lane "core"
  add_reason "no lane-specific matches -> default to core"
fi

for lane in "${LANES[@]}"; do
  printf 'lane:%s\n' "$lane"
done

for flag in "${FLAGS[@]}"; do
  printf 'flag:%s\n' "$flag"
done

for reason in "${REASONS[@]}"; do
  printf 'reason:%s\n' "$reason"
done
