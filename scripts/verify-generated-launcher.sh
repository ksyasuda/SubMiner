#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
LAUNCHER_OUT="$REPO_ROOT/dist/launcher/subminer"

if [[ ! -f "$REPO_ROOT/launcher/main.ts" ]]; then
	echo "[FAIL] launcher source missing: launcher/main.ts"
	exit 1
fi

if ! rg -n --fixed-strings -- "--outfile=\"\$(LAUNCHER_OUT)\"" "$REPO_ROOT/Makefile" >/dev/null; then
	echo "[FAIL] Makefile build-launcher target is not writing to dist/launcher/subminer"
	exit 1
fi

if [[ ! -f "$LAUNCHER_OUT" ]]; then
	echo "[FAIL] generated launcher not found at dist/launcher/subminer"
	echo "       run: make build-launcher"
	exit 1
fi

if [[ ! -x "$LAUNCHER_OUT" ]]; then
	echo "[FAIL] generated launcher is not executable: dist/launcher/subminer"
	exit 1
fi

if [[ -f "$REPO_ROOT/subminer" ]]; then
	echo "[FAIL] stale repo-root launcher artifact found: ./subminer"
	echo "       expected generated location: dist/launcher/subminer"
	echo "       remove stale artifact and use: make build-launcher"
	exit 1
fi

if git -C "$REPO_ROOT" ls-files --error-unmatch dist/launcher/subminer >/dev/null 2>&1; then
	echo "[FAIL] dist/launcher/subminer is tracked by git; generated artifacts must remain untracked"
	exit 1
fi

echo "[OK] launcher workflow verified"
echo "     source: launcher/*.ts"
echo "     generated artifact: dist/launcher/subminer"
