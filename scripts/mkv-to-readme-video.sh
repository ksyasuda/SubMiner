#!/usr/bin/env bash

set -euo pipefail

usage() {
  cat <<'EOF'
Usage:
  scripts/mkv-to-readme-video.sh <input.mkv>

Description:
  Generates two browser-friendly files next to the input file:
  - <name>.mp4  (H.264 + AAC, prefers NVIDIA GPU if available)
  - <name>.webm (AV1/VP9 + Opus, prefers NVIDIA GPU if available)
EOF
}

if [[ "${1:-}" == "-h" || "${1:-}" == "--help" ]]; then
  usage
  exit 0
fi

if [[ $# -ne 1 ]]; then
  usage
  exit 1
fi

if ! command -v ffmpeg >/dev/null 2>&1; then
  echo "Error: ffmpeg is not installed or not in PATH." >&2
  exit 1
fi

input="$1"

if [[ ! -f "$input" ]]; then
  echo "Error: input file not found: $input" >&2
  exit 1
fi

dir="$(dirname "$input")"
filename="$(basename "$input")"
base="${filename%.*}"

mp4_out="$dir/$base.mp4"
webm_out="$dir/$base.webm"

has_encoder() {
  local encoder="$1"
  ffmpeg -hide_banner -encoders 2>/dev/null | grep -qE "[[:space:]]${encoder}[[:space:]]"
}

echo "Generating MP4: $mp4_out"
if has_encoder "h264_nvenc"; then
  echo "Using GPU encoder for MP4: h264_nvenc"
  ffmpeg -y -i "$input" \
    -c:v h264_nvenc -preset p5 -cq 20 \
    -profile:v high -pix_fmt yuv420p \
    -movflags +faststart \
    -c:a aac -b:a 160k \
    "$mp4_out"
else
  echo "Using CPU encoder for MP4: libx264"
  ffmpeg -y -i "$input" \
    -c:v libx264 -preset slow -crf 20 \
    -profile:v high -level 4.1 -pix_fmt yuv420p \
    -movflags +faststart \
    -c:a aac -b:a 160k \
    "$mp4_out"
fi

echo "Generating WebM: $webm_out"
if has_encoder "av1_nvenc"; then
  echo "Using GPU encoder for WebM: av1_nvenc"
  ffmpeg -y -i "$input" \
    -c:v av1_nvenc -preset p5 -cq 30 -b:v 0 \
    -c:a libopus -b:a 128k \
    "$webm_out"
else
  echo "Using CPU encoder for WebM: libvpx-vp9"
  ffmpeg -y -i "$input" \
    -c:v libvpx-vp9 -crf 32 -b:v 0 \
    -c:a libopus -b:a 128k \
    "$webm_out"
fi

echo "Done."
echo "MP4:  $mp4_out"
echo "WebM: $webm_out"
