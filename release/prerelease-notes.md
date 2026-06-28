> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

<!-- prerelease-base-version: 0.17.1 -->

## Highlights
### Added

- **YouTube Media Cache Mode**: A new `youtube.mediaCache.mode` setting (`direct` or `background`) lets you choose how SubMiner extracts audio and image from YouTube cards.
  - In background mode, SubMiner creates a text-only card immediately, downloads a yt-dlp media cache (capped at 720p by default), and fills audio and image fields once the file is ready — with overlay and OSD notifications when the download starts and when media is available.
  - If a background download fails, SubMiner now notifies you and clears any pending media updates rather than leaving cards silently incomplete.

### Fixed

- **Log Export**: Log filenames now use your local date, so exporting logs near midnight no longer pulls stale files from the previous UTC day. Export redaction has also been expanded to mask a broader range of sensitive data, including IP addresses, email addresses, authentication and cookie headers, yt-dlp cookie arguments, URL credentials, and signed YouTube media URLs.

- **YouTube Card Media Reliability**: Direct stream extraction now uses safer ffmpeg options and skips stale or cached stream map entries to reduce failed media generation. Background cache downloads are hardened with IPv4 and extractor retry flags, stale cache files are cleaned up on startup and before new downloads, and in-flight background downloads are stopped automatically when switching back to direct mode.

## What's Changed

- feat(youtube): add mediaCache mode and safer stream media extraction by @ksyasuda in #130
- fix(logs): use local date for log filenames and expand export redaction by @ksyasuda in #131

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
