#!/bin/bash

set -euo pipefail

cat <<'EOF'
patch-yomitan.sh is retired.

SubMiner now uses the forked source submodule at vendor/subminer-yomitan and builds the
Chromium extension artifact into build/yomitan.

Use:
  git submodule update --init --recursive
  bun run build:yomitan

If you need to change Electron compatibility behavior, patch the forked source repo and rebuild.
EOF
