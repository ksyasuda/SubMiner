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
srcinfo="${pkg_dir}/.SRCINFO"

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
    found_pkgver = 0
    found_sha_block = 0
  }
  /^pkgver=/ {
    print "pkgver=" version
    found_pkgver = 1
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
      found_sha_block = 1
    }
    next
  }
  {
    print
  }
  END {
    if (!found_pkgver) {
      print "Missing pkgver= line in PKGBUILD" > "/dev/stderr"
      exit 1
    }
    if (!found_sha_block) {
      print "Missing sha256sums block in PKGBUILD" > "/dev/stderr"
      exit 1
    }
  }
  ' "$pkgbuild" > "$tmpfile"
mv "$tmpfile" "$pkgbuild"

if [[ ! -f "$srcinfo" ]]; then
  echo "Missing .SRCINFO at $srcinfo" >&2
  exit 1
fi

tmpfile="$(mktemp)"
awk \
  -v version="$version" \
  -v sum_appimage="${sha256sums[0]}" \
  -v sum_wrapper="${sha256sums[1]}" \
  -v sum_assets="${sha256sums[2]}" \
  '
  BEGIN {
    sha_index = 0
    found_pkgver = 0
    found_provides = 0
    found_noextract = 0
    found_source_appimage = 0
    found_source_wrapper = 0
    found_source_assets = 0
  }
  /^\tpkgver = / {
    print "\tpkgver = " version
    found_pkgver = 1
    next
  }
  /^\tprovides = subminer=/ {
    print "\tprovides = subminer=" version
    found_provides = 1
    next
  }
  /^\tnoextract = SubMiner-.*\.AppImage$/ {
    print "\tnoextract = SubMiner-" version ".AppImage"
    found_noextract = 1
    next
  }
  /^\tsource = SubMiner-.*\.AppImage::https:\/\/github\.com\/ksyasuda\/SubMiner\/releases\/download\/v.*\/SubMiner-.*\.AppImage$/ {
    print "\tsource = SubMiner-" version ".AppImage::https://github.com/ksyasuda/SubMiner/releases/download/v" version "/SubMiner-" version ".AppImage"
    found_source_appimage = 1
    next
  }
  /^\tsource = subminer-.*::https:\/\/github\.com\/ksyasuda\/SubMiner\/releases\/download\/v.*\/subminer$/ {
    print "\tsource = subminer-" version "::https://github.com/ksyasuda/SubMiner/releases/download/v" version "/subminer"
    found_source_wrapper = 1
    next
  }
  /^\tsource = subminer-assets-.*\.tar\.gz::https:\/\/github\.com\/ksyasuda\/SubMiner\/releases\/download\/v.*\/subminer-assets\.tar\.gz$/ {
    print "\tsource = subminer-assets-" version ".tar.gz::https://github.com/ksyasuda/SubMiner/releases/download/v" version "/subminer-assets.tar.gz"
    found_source_assets = 1
    next
  }
  /^\tsha256sums = / {
    sha_index += 1
    if (sha_index == 1) {
      print "\tsha256sums = " sum_appimage
      next
    }
    if (sha_index == 2) {
      print "\tsha256sums = " sum_wrapper
      next
    }
    if (sha_index == 3) {
      print "\tsha256sums = " sum_assets
      next
    }
  }
  {
    print
  }
  END {
    if (!found_pkgver) {
      print "Missing pkgver entry in .SRCINFO" > "/dev/stderr"
      exit 1
    }
    if (!found_provides) {
      print "Missing provides entry in .SRCINFO" > "/dev/stderr"
      exit 1
    }
    if (!found_noextract) {
      print "Missing noextract entry in .SRCINFO" > "/dev/stderr"
      exit 1
    }
    if (!found_source_appimage || !found_source_wrapper || !found_source_assets) {
      print "Missing source entry in .SRCINFO" > "/dev/stderr"
      exit 1
    }
    if (sha_index < 3) {
      print "Missing sha256sums entries in .SRCINFO" > "/dev/stderr"
      exit 1
    }
  }
  ' "$srcinfo" > "$tmpfile"
mv "$tmpfile" "$srcinfo"
