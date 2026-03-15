#!/usr/bin/env bash
set -euo pipefail

usage() {
  cat <<'EOF'
Usage: scripts/update-aur-package.sh --pkg-dir <dir> --version <version> --appimage <path> --wrapper <path> --assets <path>
EOF
}

pkg_dir=
version=
appimage=
wrapper=
assets=

while [[ $# -gt 0 ]]; do
  case "$1" in
    --pkg-dir)
      pkg_dir="${2:-}"
      shift 2
      ;;
    --version)
      version="${2:-}"
      shift 2
      ;;
    --appimage)
      appimage="${2:-}"
      shift 2
      ;;
    --wrapper)
      wrapper="${2:-}"
      shift 2
      ;;
    --assets)
      assets="${2:-}"
      shift 2
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      echo "Unknown argument: $1" >&2
      usage >&2
      exit 1
      ;;
  esac
done

if [[ -z "$pkg_dir" || -z "$version" || -z "$appimage" || -z "$wrapper" || -z "$assets" ]]; then
  usage >&2
  exit 1
fi

version="${version#v}"
pkgbuild="${pkg_dir}/PKGBUILD"

if [[ ! -f "$pkgbuild" ]]; then
  echo "Missing PKGBUILD at $pkgbuild" >&2
  exit 1
fi

for artifact in "$appimage" "$wrapper" "$assets"; do
  if [[ ! -f "$artifact" ]]; then
    echo "Missing artifact: $artifact" >&2
    exit 1
  fi
done

mapfile -t sha256sums < <(sha256sum "$appimage" "$wrapper" "$assets" | awk '{print $1}')

tmpfile="$(mktemp)"
awk \
  -v version="$version" \
  -v sum_appimage="${sha256sums[0]}" \
  -v sum_wrapper="${sha256sums[1]}" \
  -v sum_assets="${sha256sums[2]}" \
  '
  BEGIN {
    in_sha_block = 0
  }
  /^pkgver=/ {
    print "pkgver=" version
    next
  }
  /^sha256sums=\(/ {
    print "sha256sums=("
    print "\047" sum_appimage "\047"
    print "\047" sum_wrapper "\047"
    print "\047" sum_assets "\047"
    in_sha_block = 1
    next
  }
  in_sha_block {
    if ($0 ~ /^\)/) {
      print ")"
      in_sha_block = 0
    }
    next
  }
  {
    print
  }
  ' "$pkgbuild" > "$tmpfile"
mv "$tmpfile" "$pkgbuild"

(
  cd "$pkg_dir"
  makepkg --printsrcinfo > .SRCINFO
)
