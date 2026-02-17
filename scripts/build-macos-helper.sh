#!/bin/bash
# Build macOS window tracking helper binary

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SWIFT_SOURCE="$SCRIPT_DIR/get-mpv-window-macos.swift"
OUTPUT_DIR="$SCRIPT_DIR/../dist/scripts"
OUTPUT_BINARY="$OUTPUT_DIR/get-mpv-window-macos"
OUTPUT_SOURCE_COPY="$OUTPUT_DIR/get-mpv-window-macos.swift"

fallback_to_source() {
  echo "Falling back to source fallback: $OUTPUT_SOURCE_COPY"
  mkdir -p "$OUTPUT_DIR"
  cp "$SWIFT_SOURCE" "$OUTPUT_SOURCE_COPY"
}

build_swift_helper() {
  echo "Compiling macOS window tracking helper..."
  if ! command -v swiftc >/dev/null 2>&1; then
    echo "swiftc not found in PATH; skipping compilation."
    return 1
  fi

  if ! swiftc -O "$SWIFT_SOURCE" -o "$OUTPUT_BINARY"; then
    return 1
  fi

  chmod +x "$OUTPUT_BINARY"
  echo "✓ Built $OUTPUT_BINARY"
  return 0
}

# Optional skip flag for non-macOS CI/dev environments
if [[ "${SUBMINER_SKIP_MACOS_HELPER_BUILD:-}" == "1" ]]; then
  echo "Skipping macOS helper build (SUBMINER_SKIP_MACOS_HELPER_BUILD=1)"
  fallback_to_source
  exit 0
fi

# Only build on macOS
if [[ "$(uname)" != "Darwin" ]]; then
  echo "Skipping macOS helper build (not on macOS)"
  fallback_to_source
  exit 0
fi

# Create output directory
mkdir -p "$OUTPUT_DIR"

# Compile Swift script to binary, fallback to source if unavailable or compilation fails
if ! build_swift_helper; then
  fallback_to_source
fi
