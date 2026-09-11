#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
LAUNCHER_DIR="$REPO_ROOT/dist/launcher"
LAUNCHER_OUT="$LAUNCHER_DIR/subminer"
EXPECTED_ARTIFACTS=(prepare.cjs subminer subminer.cmd subminer.js version)

if [[ ! -f "$REPO_ROOT/launcher/main.ts" ]]; then
	echo "[FAIL] launcher source missing: launcher/main.ts"
	exit 1
fi

if ! grep -F -- "bun run build:launcher" "$REPO_ROOT/Makefile" >/dev/null; then
	echo "[FAIL] Makefile build-launcher target does not call the canonical package script"
	exit 1
fi

for artifact in "${EXPECTED_ARTIFACTS[@]}"; do
	if [[ ! -f "$LAUNCHER_DIR/$artifact" ]]; then
		echo "[FAIL] generated launcher artifact missing: dist/launcher/$artifact"
		echo "       run: make build-launcher"
		exit 1
	fi
done

for artifact_path in "$LAUNCHER_DIR"/*; do
	artifact="${artifact_path##*/}"
	case "$artifact" in
		prepare.cjs | subminer | subminer.cmd | subminer.js | version) ;;
		*)
			echo "[FAIL] dist/launcher contains an unexpected runtime artifact: $artifact"
			exit 1
			;;
	esac
done

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

if git -C "$REPO_ROOT" ls-files --error-unmatch dist/launcher >/dev/null 2>&1; then
	echo "[FAIL] dist/launcher contains tracked files; generated artifacts must remain untracked"
	exit 1
fi

echo "[OK] launcher workflow verified"
echo "     source: launcher/*.ts"
echo "     generated artifacts: ${EXPECTED_ARTIFACTS[*]}"
