#!/bin/bash
#
# SubMiner - All-in-one sentence mining overlay
# Copyright (C) 2024 sudacode
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
# patch-texthooker.sh - Apply patches to texthooker-ui
#
# This script patches texthooker-ui to handle empty sentences from mpv.
# When subtitles disappear, mpv sends {"sentence":""} which would otherwise
# display as raw JSON on the texthooker page.
#
# Usage: ./patch-texthooker.sh [texthooker_dir]
#   texthooker_dir: Path to the texthooker-ui directory (default: vendor/texthooker-ui)
#

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TEXTHOOKER_DIR="${1:-$SCRIPT_DIR/../vendor/texthooker-ui}"

if [ ! -d "$TEXTHOOKER_DIR" ]; then
    echo "Error: texthooker-ui directory not found: $TEXTHOOKER_DIR"
    exit 1
fi

echo "Patching texthooker-ui in: $TEXTHOOKER_DIR"

SOCKET_TS="$TEXTHOOKER_DIR/src/socket.ts"

if [ ! -f "$SOCKET_TS" ]; then
    echo "Error: socket.ts not found at $SOCKET_TS"
    exit 1
fi

echo "Patching socket.ts..."

# Patch 1: Change || to ?? (nullish coalescing)
# This ensures empty string is kept instead of falling back to raw JSON
if grep -q '\.sentence ?? event\.data' "$SOCKET_TS"; then
    echo "  - Nullish coalescing already patched, skipping"
else
    sed -i 's/\.sentence || event\.data/.sentence ?? event.data/' "$SOCKET_TS"
    echo "  - Changed || to ?? (nullish coalescing)"
fi

# Patch 2: Skip emitting empty lines
# This prevents empty sentences from being added to the UI
if grep -q "if (line)" "$SOCKET_TS"; then
    echo "  - Empty line check already patched, skipping"
else
    sed -i 's/\t\tnewLine\$\.next(\[line, LineType\.SOCKET\]);/\t\tif (line) {\n\t\t\tnewLine$.next([line, LineType.SOCKET]);\n\t\t}/' "$SOCKET_TS"
    echo "  - Added empty line check"
fi

echo ""
echo "texthooker-ui patching complete!"
echo ""
echo "Changes applied:"
echo "  1. socket.ts: Use ?? instead of || to preserve empty strings"
echo "  2. socket.ts: Skip emitting empty sentences"
echo ""
echo "To rebuild: cd vendor/texthooker-ui && pnpm build"
