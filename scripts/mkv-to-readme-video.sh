#!/usr/bin/env bash

set -euo pipefail

usage() {
	cat << 'USAGE'
Usage:
  scripts/mkv-to-readme-video.sh [--force] <input.mkv>

Description:
  Generates two browser-friendly files next to the input file:
  - <name>.mp4  (H.264 + AAC, prefers NVIDIA GPU if available)
  - <name>.webm (AV1/VP9 + Opus, prefers NVIDIA GPU if available)
  - <name>-poster.jpg (single frame for video poster fallback)
  - <name>.webp (animated, only when --webp is provided)

Options:
  -f, --force  Overwrite existing output files
  -w, --webp   Generate animated WebP preview

Encoding profile:
  - Crop: 1920x1080 at x=760 y=200
  - MP4:  H.264 + AAC
  - WebM: AV1/VP9 + Opus at 30 fps
USAGE
}

force=0
generate_webp=0
input=""

while [[ $# -gt 0 ]]; do
	case "$1" in
		-h | --help)
			usage
			exit 0
			;;
		-f | --force)
			force=1
			;;
		-w | --webp)
			generate_webp=1
			;;
		-*)
			echo "Error: unknown option: $1" >&2
			usage
			exit 1
			;;
		*)
			if [[ -n "$input" ]]; then
				echo "Error: expected exactly one input file." >&2
				usage
				exit 1
			fi
			input="$1"
			;;
	esac
	shift
done

if [[ -z "$input" ]]; then
	usage
	exit 1
fi

if ! command -v ffmpeg > /dev/null 2>&1; then
	echo "Error: ffmpeg is not installed or not in PATH." >&2
	exit 1
fi

if [[ ! -f "$input" ]]; then
	echo "Error: input file not found: $input" >&2
	exit 1
fi

dir="$(dirname "$input")"
filename="$(basename "$input")"
base="${filename%.*}"

mp4_out="$dir/$base.mp4"
webm_out="$dir/$base.webm"
webp_out="$dir/$base.webp"
poster_out="$dir/$base-poster.jpg"

overwrite_flag="-n"
if [[ "$force" -eq 1 ]]; then
	overwrite_flag="-y"
fi

if [[ "$force" -eq 0 ]]; then
	outputs=("$mp4_out" "$webm_out" "$poster_out")
	if [[ "$generate_webp" -eq 1 ]]; then
		outputs+=("$webp_out")
	fi
	for output in "${outputs[@]}"; do
		if [[ -e "$output" ]]; then
			echo "Error: output exists: $output (use --force to overwrite)" >&2
			exit 1
		fi
	done
fi

has_encoder() {
	local encoder="$1"
	ffmpeg -hide_banner -encoders 2> /dev/null | awk -v encoder="$encoder" '$2 == encoder { found = 1 } END { exit(found ? 0 : 1) }'
}

pick_webp_encoder() {
	if has_encoder "libwebp_anim"; then
		echo "libwebp_anim"
		return 0
	fi
	if has_encoder "libwebp"; then
		echo "libwebp"
		return 0
	fi
	return 1
}

crop_vf="crop=1920:1080:760:205"
webm_vf="${crop_vf},fps=30"

echo "Generating MP4: $mp4_out"
if has_encoder "h264_nvenc"; then
	echo "Trying GPU encoder for MP4: h264_nvenc"
	if ffmpeg "$overwrite_flag" -i "$input" \
		-vf "$crop_vf" \
		-c:v h264_nvenc -preset p6 -rc:v vbr -cq:v 20 -b:v 0 \
		-pix_fmt yuv420p -movflags +faststart \
		-c:a aac -b:a 160k \
		"$mp4_out"; then
		:
	else
		echo "GPU MP4 encode failed; retrying with CPU encoder: libx264"
		ffmpeg "$overwrite_flag" -i "$input" \
			-vf "$crop_vf" \
			-c:v libx264 -preset slow -crf 20 \
			-profile:v high -level 4.1 -pix_fmt yuv420p \
			-movflags +faststart \
			-c:a aac -b:a 160k \
			"$mp4_out"
	fi
else
	echo "Using CPU encoder for MP4: libx264"
	ffmpeg "$overwrite_flag" -i "$input" \
		-vf "$crop_vf" \
		-c:v libx264 -preset slow -crf 20 \
		-profile:v high -level 4.1 -pix_fmt yuv420p \
		-movflags +faststart \
		-c:a aac -b:a 160k \
		"$mp4_out"
fi

echo "Generating WebM: $webm_out"
if has_encoder "av1_nvenc"; then
	echo "Trying GPU encoder for WebM: av1_nvenc"
	if ffmpeg "$overwrite_flag" -i "$input" \
		-vf "$webm_vf" \
		-c:v av1_nvenc -preset p6 -cq:v 34 -b:v 0 \
		-c:a libopus -b:a 96k \
		"$webm_out"; then
		:
	else
		echo "GPU WebM encode failed; retrying with CPU encoder: libvpx-vp9"
		ffmpeg "$overwrite_flag" -i "$input" \
			-vf "$webm_vf" \
			-c:v libvpx-vp9 -crf 34 -b:v 0 \
			-row-mt 1 -threads 8 \
			-c:a libopus -b:a 96k \
			"$webm_out"
	fi
else
	echo "Using CPU encoder for WebM: libvpx-vp9"
	ffmpeg "$overwrite_flag" -i "$input" \
		-vf "$webm_vf" \
		-c:v libvpx-vp9 -crf 34 -b:v 0 \
		-row-mt 1 -threads 8 \
		-c:a libopus -b:a 96k \
		"$webm_out"
fi

if [[ "$generate_webp" -eq 1 ]]; then
	if ! webp_encoder="$(pick_webp_encoder)"; then
		echo "Error: encoder not found: libwebp" >&2
		exit 1
	fi
	echo "Generating animated WebP with $webp_encoder: $webp_out"
	ffmpeg "$overwrite_flag" -i "$input" \
		-vf "${crop_vf},fps=24,scale=960:-1:flags=lanczos" \
		-c:v "$webp_encoder" \
		-q:v 80 \
		-loop 0 \
		-an \
		"$webp_out"
fi

echo "Generating poster: $poster_out"
ffmpeg "$overwrite_flag" -ss 00:00:05 -i "$input" \
	-vf "$crop_vf" \
	-vframes 1 \
	-q:v 2 \
	"$poster_out"

echo "Done."
echo "MP4:  $mp4_out"
echo "WebM: $webm_out"
if [[ "$generate_webp" -eq 1 ]]; then
	echo "WebP:  $webp_out"
fi
echo "Poster: $poster_out"
