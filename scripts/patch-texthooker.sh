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
# This script patches texthooker-ui to:
# 1) Handle empty sentences from mpv.
# 2) Apply SubMiner default texthooker styling.
#
# Usage: ./patch-texthooker.sh [texthooker_dir]
#   texthooker_dir: Path to the texthooker-ui directory (default: vendor/texthooker-ui)
#

set -e

sed_in_place() {
    local script="$1"
    local file="$2"

    if sed -i '' "$script" "$file" 2>/dev/null; then
        return 0
    fi

    sed -i "$script" "$file"
}

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TEXTHOOKER_DIR="${1:-$SCRIPT_DIR/../vendor/texthooker-ui}"

if [ ! -d "$TEXTHOOKER_DIR" ]; then
    echo "Error: texthooker-ui directory not found: $TEXTHOOKER_DIR"
    exit 1
fi

echo "Patching texthooker-ui in: $TEXTHOOKER_DIR"

SOCKET_TS="$TEXTHOOKER_DIR/src/socket.ts"
APP_CSS="$TEXTHOOKER_DIR/src/app.css"
STORES_TS="$TEXTHOOKER_DIR/src/stores/stores.ts"

if [ ! -f "$SOCKET_TS" ]; then
    echo "Error: socket.ts not found at $SOCKET_TS"
    exit 1
fi

if [ ! -f "$APP_CSS" ]; then
    echo "Error: app.css not found at $APP_CSS"
    exit 1
fi

if [ ! -f "$STORES_TS" ]; then
    echo "Error: stores.ts not found at $STORES_TS"
    exit 1
fi

echo "Patching socket.ts..."

# Patch 1: Change || to ?? (nullish coalescing)
# This ensures empty string is kept instead of falling back to raw JSON
if grep -q '\.sentence ?? event\.data' "$SOCKET_TS"; then
    echo "  - Nullish coalescing already patched, skipping"
else
    sed_in_place 's/\.sentence || event\.data/.sentence ?? event.data/' "$SOCKET_TS"
    echo "  - Changed || to ?? (nullish coalescing)"
fi

# Patch 2: Skip emitting empty lines
# This prevents empty sentences from being added to the UI
if grep -q "if (line)" "$SOCKET_TS"; then
    echo "  - Empty line check already patched, skipping"
else
    sed_in_place 's/\t\tnewLine\$\.next(\[line, LineType\.SOCKET\]);/\t\tif (line) {\n\t\t\tnewLine$.next([line, LineType.SOCKET]);\n\t\t}/' "$SOCKET_TS"
    echo "  - Added empty line check"
fi

echo "Patching app.css..."

# Patch 3: Apply SubMiner default texthooker-ui styling
if grep -q "SUBMINER_DEFAULT_STYLE_START" "$APP_CSS"; then
    echo "  - Default SubMiner CSS already patched, skipping"
else
    cat >> "$APP_CSS" << 'EOF'

/* SUBMINER_DEFAULT_STYLE_START */
:root {
    --sm-bg: #1a1b2e;
    --sm-surface: #222436;
    --sm-border: rgba(255, 255, 255, 0.05);
    --sm-text: #c8d3f5;
    --sm-text-muted: #636da6;
    --sm-hover-bg: rgba(130, 170, 255, 0.06);
    --sm-scrollbar: rgba(255, 255, 255, 0.08);
    --sm-scrollbar-hover: rgba(255, 255, 255, 0.15);
}

html,
body {
    margin: 0;
    display: flex;
    flex: 1;
    min-height: 100%;
    overflow: auto;
    background: var(--sm-bg);
    color: var(--sm-text);
}

body[data-theme] {
    background: var(--sm-bg);
    color: var(--sm-text);
}

main,
header,
#pip-container {
    background: transparent;
    color: var(--sm-text);
}

header.bg-base-100,
header.bg-base-200 {
    background: transparent;
}

main,
main.flex {
    font-family: 'Noto Sans CJK JP', 'Hiragino Sans', system-ui, sans-serif;
    padding: 0.5rem min(4vw, 2rem);
    line-height: 1.7;
    gap: 0;
}

p,
p.cursor-pointer {
    font-family: 'Noto Sans CJK JP', 'Hiragino Sans', system-ui, sans-serif;
    font-size: clamp(18px, 2vw, 26px);
    letter-spacing: 0.04em;
    line-height: 1.65;
    white-space: normal;
    margin: 0;
    padding: 0.65rem 1rem;
    border: none;
    border-bottom: 1px solid var(--sm-border);
    border-radius: 0;
    background: transparent;
    transition: background 0.2s ease;
}

p:hover,
p.cursor-pointer:hover {
    background: var(--sm-hover-bg);
}

p:last-child,
p.cursor-pointer:last-child {
    border-bottom: none;
}

p.cursor-pointer.whitespace-pre-wrap {
    white-space: normal;
}

p.cursor-pointer.my-2 {
    margin: 0;
}

p.cursor-pointer.py-4,
p.cursor-pointer.py-2,
p.cursor-pointer.px-2,
p.cursor-pointer.px-4 {
    padding: 0.65rem 1rem;
}

*::-webkit-scrollbar {
    width: 4px;
}

*::-webkit-scrollbar-track {
    background: transparent;
}

*::-webkit-scrollbar-thumb {
    background: var(--sm-scrollbar);
    border-radius: 2px;
}

*::-webkit-scrollbar-thumb:hover {
    background: var(--sm-scrollbar-hover);
}
/* SUBMINER_DEFAULT_STYLE_END */
EOF
    echo "  - Added default SubMiner CSS block"
fi

echo "Patching stores.ts defaults..."

# Patch 4: Change default settings for title/whitespace/animation/reconnect
if grep -q "preserveWhitespace\\$: false" "$STORES_TS" && \
   grep -q "removeAllWhitespace\\$: true" "$STORES_TS" && \
   grep -q "enableLineAnimation\\$: true" "$STORES_TS" && \
   grep -q "continuousReconnect\\$: true" "$STORES_TS" && \
   grep -q "windowTitle\\$: 'SubMiner Texthooker'" "$STORES_TS"; then
    echo "  - Default settings already patched, skipping"
else
    sed_in_place "s/windowTitle\\$: '',/windowTitle\\$: 'SubMiner Texthooker',/" "$STORES_TS"
    sed_in_place 's/preserveWhitespace\$: true,/preserveWhitespace\$: false,/' "$STORES_TS"
    sed_in_place 's/removeAllWhitespace\$: false,/removeAllWhitespace\$: true,/' "$STORES_TS"
    sed_in_place 's/enableLineAnimation\$: false,/enableLineAnimation\$: true,/' "$STORES_TS"
    sed_in_place 's/continuousReconnect\$: false,/continuousReconnect\$: true,/' "$STORES_TS"
    echo "  - Updated default settings (title/whitespace/animation/reconnect)"
fi

echo ""
echo "texthooker-ui patching complete!"
echo ""
echo "Changes applied:"
echo "  1. socket.ts: Use ?? instead of || to preserve empty strings"
echo "  2. socket.ts: Skip emitting empty sentences"
echo "  3. app.css: Apply SubMiner default styling (without !important)"
echo "  4. stores.ts: Update default settings for title/whitespace/animation/reconnect"
echo ""
echo "To rebuild: cd vendor/texthooker-ui && bun run build"
