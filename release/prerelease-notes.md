> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

<!-- prerelease-base-version: 0.19.4 -->

## Highlights
### Added
- **Library Merge and Move**
  - Duplicate library cards for the same show can now be combined: select cards in the library grid, choose "Merge Selected," and pick which entry to keep. Sessions, mined cards, and watch time move over, and future episodes stay matched to the merged card.
  - Episodes can be reassigned to a different library entry with a "→" button on the episode row, fixing cases where a file lands under the wrong title. Manual assignments survive later filename parsing, Jellyfin refreshes, and season repair.
  - Exact AniList matches with compatible seasons now merge automatically, while fuzzy matches show up as a dismissible "Possible duplicate" prompt instead of merging without confirmation.

### Fixed
- **Anki Audio Generation on Network Drives**
  - Fixed sentence-audio generation timing out on slow network-mounted video files with many subtitle and font-attachment streams.
  - Extraction now uses bounded probing and a two-minute budget, and failures show a clear error instead of a cryptic one.
- **Duplicate Subtitle Line Stats**
  - Fixed karaoke openings and animated signs (which record one subtitle event per animation frame) inflating word and kanji counts and skewing "Top Repeated Words." Ordinary repeated dialogue and rewatches are unaffected.
  - Already-inflated stats can be cleaned up with the new "Duplicates" button in the Vocabulary tab, or `subminer stats cleanup --duplicate-lines` (supports `--dry-run` and `--lookback-days`). Only the affected subtitle lines and vocabulary counts are touched; watch time and lines-seen totals are untouched.
- **Overlay Modals on macOS and Windows**
  - Fixed overlay modals and the stats window opening on the wrong macOS Space, or forcing a Space switch, when mpv is fullscreen. They now open above fullscreen mpv on its current Space.
  - Modals are now prewarmed on macOS and Windows so shortcuts open them promptly, and Windows keeps the hidden modal responsive between sessions.
- **Wayland File Drag-and-Drop**
  - Fixed dragging subtitle and video files from file managers like Thunar onto the overlay on native Wayland; dropped files are now resolved and sent to mpv.
- **Windows Mouse Lag**
  - Fixed system-wide mouse lag while SubMiner is running on Windows, caused by a global mouse hook and blocking window lookups during click-through tracking.
- **Mining Clip Accuracy**
  - Fixed mined audio and animated image clips sometimes capturing the wrong subtitle line when audio extraction was slow. The clip range is now locked in at the moment of lookup, so audio and image clips always match.
- **Linux Notifications**
  - Character dictionary progress notifications on Linux now update in place instead of flickering off and back on with every status change.
- **Stats Delete Performance**
  - Fixed stats deletes freezing the dashboard; deletes now reliably run off the main thread, with automatic retry if the delete worker crashes.
  - Deletes, library merges/moves, and AniList reassignments are now much faster because totals are updated incrementally instead of rebuilt from scratch, and no longer erase lifetime totals older than the recent session history.
  - Session deletes on large libraries dropped from minutes to milliseconds.

### Docs
- **Feature Demos Page**
  - Hidden the unfinished feature demos page from the documentation sidebar; it's still reachable by direct URL.

## What's Changed

- feat(stats): add library entry merge and episode move by @ksyasuda in #190
- fix(stats): stop counting duplicate typeset subtitle lines by @ksyasuda in #191
- fix(media): tolerate slow MKV audio extraction by @ksyasuda in #195
- fix(stats): subtract lifetime totals incrementally on delete by @ksyasuda in #196
- fix(anki): snapshot mining media clip timing by @ksyasuda in #197
- fix(notifications): replace Linux progress updates in place by @ksyasuda in #198
- fix(overlay): support native Wayland file drag-and-drop by @ksyasuda in #199
- fix(overlay): keep macOS modal windows on fullscreen Spaces by @ksyasuda in #200
- fix(overlay): prevent Windows mouse lag during click-through tracking by @ksyasuda in #201

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
