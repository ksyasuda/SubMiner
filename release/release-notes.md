## Highlights

### Fixed

- **YouTube Background Cache:** Fixed Windows background media cache startup for YouTube URLs opened directly in mpv.
  - Resolved stream URLs are now tracked even when mpv still exposes the original YouTube playlist entry.
  - Queued Anki media updates can append audio and images after the cache finishes instead of staying text-only.

- **YouTube Subtitle Picker Notifications:** Manual subtitle picker requests now show immediate status while SubMiner probes tracks and opens the modal.
  - Subtitle download progress is replaced with a transient success notification after tracks load.

## What's Changed

- feat(youtube): notify on manual picker open and show success after track load by @ksyasuda in #133
- fix(youtube): recover source URL for background media cache on direct mpv open by @ksyasuda in #132

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
