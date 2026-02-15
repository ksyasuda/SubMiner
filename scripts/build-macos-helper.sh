#!/bin/bash
# Build macOS window tracking helper binary

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SWIFT_SOURCE="$SCRIPT_DIR/get-mpv-window-macos.swift"
OUTPUT_DIR="$SCRIPT_DIR/../dist/scripts"
OUTPUT_BINARY="$OUTPUT_DIR/get-mpv-window-macos"

# Only build on macOS
if [[ "$(uname)" != "Darwin" ]]; then
  echo "Skipping macOS helper build (not on macOS)"
  exit 0
fi

# Create output directory
mkdir -p "$OUTPUT_DIR"

# Compile Swift script to binary
echo "Compiling macOS window tracking helper..."
swiftc -O "$SWIFT_SOURCE" -o "$OUTPUT_BINARY"

# Make executable
chmod +x "$OUTPUT_BINARY"

echo "✓ Built $OUTPUT_BINARY"
