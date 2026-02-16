#!/usr/bin/env bash
set -euo pipefail

usage() {
	cat <<'EOF'
Usage:
  ./scripts/check-main-lines.sh [target-lines] [file]
  ./scripts/check-main-lines.sh --target <target-lines> --file <path>

  target-lines default: 1500
  file default: src/main.ts
EOF
}

target="1500"
file="src/main.ts"

if (($# == 1)) && [[ "$1" == "-h" || "$1" == "--help" ]]; then
	usage
	exit 0
fi

if [[ $# -ge 1 && "$1" != --* ]]; then
	target="$1"
	if [[ $# -ge 2 ]]; then
		file="$2"
	fi
	shift $(($# > 1 ? 2 : 1))
fi

while [[ $# -gt 0 ]]; do
	case "$1" in
	--target)
		target="$2"
		shift 2
		;;
	--file)
		file="$2"
		shift 2
		;;
	--help)
		usage
		exit 0
		;;
	*)
		echo "[ERROR] Unknown argument: $1" >&2
		usage
		exit 1
		;;
	esac
done

if [[ ! -f "$file" ]]; then
	echo "[ERROR] File not found: $file" >&2
	exit 1
fi

if ! [[ "$target" =~ ^[0-9]+$ ]]; then
	echo "[ERROR] Target line count must be an integer. Got: $target" >&2
	exit 1
fi

actual="$(wc -l <"$file" | tr -d ' ')"

echo "[INFO] $file lines: $actual (target: <= $target)"
if ((actual > target)); then
	echo "[ERROR] Line gate failed: $actual > $target" >&2
	exit 1
fi

echo "[OK] Line gate passed"
