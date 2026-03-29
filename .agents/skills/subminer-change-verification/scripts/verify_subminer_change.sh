#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR=$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)
REPO_ROOT=$(cd "$SCRIPT_DIR/../../../.." && pwd)
TARGET="$REPO_ROOT/plugins/subminer-workflow/skills/subminer-change-verification/scripts/verify_subminer_change.sh"

if [[ ! -x "$TARGET" ]]; then
  echo "Missing canonical script: $TARGET" >&2
  exit 1
fi

exec "$TARGET" "$@"
