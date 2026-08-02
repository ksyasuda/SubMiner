## Highlights
### Added
- Kiku/Lapis Word Card Type Setting
  - A new setting (Mining/Anki > Kiku/Lapis Features > "Word Card Type") lets you choose which card-type flag gets marked on Kiku/Lapis word cards, including a click-card option that SubMiner couldn't set before.
  - Handy if you only want click cards flagged instead of the default word-and-sentence marking.
  - Choosing a card type now clears any other flags automatically, so a note can't end up marked as two types at once.

### Fixed
- Yomitan Popup on macOS
  - Fixed the popup going unresponsive after mining a card — clicks outside it no longer leak through to mpv, and the overlay no longer flickers hidden and shown.
  - Scrolling over the popup now scrolls its definitions instead of seeking the video.
- YouTube Playlist Links
  - Opening a video from a playlist URL (like a Watch Later link with `list=`/`index=`) no longer times out while loading subtitles, metadata, or playback info.

## What's Changed

- feat(anki): add configurable word card type for Kiku/Lapis by @ksyasuda in #175
- fix(overlay): keep Yomitan popup interactive on macOS/Windows by @ksyasuda in #177
- fix(youtube): prevent playlist URLs from stalling yt-dlp probes by @ksyasuda in #180

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
