## Highlights
### Added

- **YouTube Media Cache Mode:** New `youtube.mediaCache.mode` setting lets you choose between `direct` (the existing behavior) and `background` modes for extracting audio and images when creating cards from YouTube videos.
  - Background mode uses a separate yt-dlp download to cache the video file locally, which can rescue media extraction when direct stream URLs are unreliable or expire quickly.
  - While the cache downloads, SubMiner creates a text-only card immediately, then fills in audio and image once the file is ready — no need to re-mine the sentence.
  - Cache downloads are capped at 720p by default (`youtube.mediaCache.maxHeight`) to keep file sizes reasonable.

### Fixed

- **Log Export:** Log filenames now use your local date, so exporting logs around UTC midnight correctly captures today's file instead of yesterday's.
  - Log exports also now redact a broader set of sensitive data — IP addresses, email addresses, auth and cookie headers, yt-dlp cookie arguments, URL credentials, API keys, and signed YouTube media URLs.

- **YouTube Card Media Generation:** Extraction is more reliable and resilient across a wider range of stream conditions.
  - Stale or cached stream maps are skipped automatically, and ffmpeg requests use safer options for resolved streams to reduce failures.
  - Background cache downloads retry on network issues (IPv4 and extractor fallbacks), notify you on failure instead of silently hanging, and clean up stale cache files on startup and before each new download.
  - Switching back to direct mode now cancels any in-progress background downloads cleanly.

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
