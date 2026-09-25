#!/usr/bin/env bash

set -euo pipefail

FILE="${1:-}"

if [[ ! -f "$FILE" ]]; then
	printf 'Not a file: %s\n' "${FILE:-<missing>}" >&2
	exit 1
fi

if ! mpv --no-config --no-terminal --msg-level=all=no --vo=null --ao=null --frames=1 -- "$FILE"; then
	printf 'Not playable by mpv: %s\n' "$FILE" >&2
	exit 1
fi

exec subminer app --dev --launch-mpv "$FILE"
