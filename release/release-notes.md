## Highlights
### Added
- **Library Duplicate & Misfiled Episode Tools**
  - Merge duplicate show cards from the Library grid: select cards and use "Merge Selected" to combine sessions, mined cards, and watch time onto one entry while keeping remembered title aliases.
  - Reassign a misfiled episode to the correct show with the "→" button on an episode row; the fix survives later filename parsing, Jellyfin refreshes, and season repair.
  - Exact AniList matches merge automatically, while likely (fuzzy) matches surface as a dismissible "Possible duplicate" suggestion instead of merging without confirmation.
- **Stats Duplicate-Line Cleanup Tool**
  - The Vocabulary tab's new Duplicates button scans a chosen time window for old karaoke/animation duplicate bursts and collapses each one to a single line after you confirm, without touching watch time or lines-seen totals.
  - The same cleanup is available from the terminal via `subminer stats cleanup --duplicate-lines`, with `--dry-run` and `--lookback-days` options.

### Changed
- **Prerelease Notes "Changes Since" Section**
  - Prerelease release notes now open with a "Changes since" section listing only what changed since the previous beta/RC of the same version, shown above the full cumulative highlights.

### Fixed
- **Subtitle Deduplication & Karaoke Reconstruction**
  - Typeset ASS karaoke and animated signs are reconstructed into their authored line and shown once, instead of flooding the overlay, subtitle sidebar, immersion history, sentence mining, and stats with per-frame glyph fragments and repeated lyric bursts (a lyric could previously pin itself to the top of "Top Repeated Words").
  - The same deduplication now applies consistently everywhere, including embedded subtitles extracted from network-mounted (SMB/NFS) media and the secondary subtitle overlay, while ordinary repeated dialogue, signs, and rewatches remain unaffected.
  - Secondary subtitle overlays no longer clip long lines after about four rows, and no longer show scattered-letter or duplicated text while embedded subtitles are still being extracted.
- **Character Dictionary Reliability & Notifications**
  - Character dictionary generation, rebuilds, and imports no longer freeze the app or trigger "not responding" dialogs on large dictionaries; the heavy work now runs off the main UI thread.
  - Dictionaries are reused instead of being regenerated on every launch when no name splits were found, and portraits reappear correctly once the cached portrait index finishes loading.
  - Linux desktop progress notifications, including on AppImage installs, now update in place instead of flickering closed and reopening.
- **Overlay Startup Reliability**
  - The overlay no longer gets stuck on an endless "Overlay loading" screen when mpv's connection stalls at startup; connections now time out and retry, and a clear error appears if content still isn't ready after 30 seconds.
- **Overlay Modal Windows (macOS & Windows)**
  - Modal windows such as Settings prewarm so shortcuts open them promptly on first press.
  - On Windows, the hidden modal renderer now refreshes between sessions so later modals stay interactive.
  - On macOS, reused modals and the stats window open above fullscreen mpv on the correct Space instead of jumping to another desktop; the overlay-attach helper also now supports macOS 12.0+, fixing "Overlay loading" getting stuck on older macOS versions.
- **Windows Mouse Lag**
  - Fixed system-wide mouse lag while SubMiner is running: the overlay no longer installs a global mouse hook, and the mpv window tracker no longer blocks the app with repeated command-line lookups.
- **Linux Overlay & Launcher Fixes**
  - Native Wayland drag-and-drop from file managers such as Thunar now works, so subtitle and video files dropped on the overlay reach mpv.
  - Fixed missing MKV thumbnails in the rofi file picker on systems that only advertise legacy Matroska MIME aliases.
- **Sentence Mining Audio & Clip Accuracy**
  - Sentence-audio generation no longer times out on slow network-mounted MKV files with many subtitle/font streams; probing is now bounded with a two-minute extraction budget and a clear error instead of a raw failure.
  - Mined audio and animated clips now capture the exact subtitle line that was mined, instead of whatever line was on screen after audio extraction finished, fixing too-short or misaligned clips.
- **Stats Reliability & Performance**
  - Fixed transient database-lock errors when multiple stats workers wrote at once.
  - Stats deletes, library merges, video moves, and AniList reassignments no longer freeze the dashboard or rebuild lifetime totals from scratch, so they're fast and preserve lifetime totals older than the recent session-retention window; session deletes on large databases dropped from minutes to milliseconds.
- **Vocabulary Tab Accuracy**
  - Vocabulary totals and charts now count all tracked vocabulary instead of only the first page, with new-word history rebuilt from corrected daily rollups to match.
  - Calendar charts keep the correct local date in time zones west of UTC, and vocabulary cards/charts now refresh automatically and retry after the word exclusion list changes.

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
- fix(stats): report complete vocabulary totals and new-word history by @ksyasuda in #202
- fix(mpv): recover from stalled IPC connects by @ksyasuda in #204
- fix(dictionary): prevent freezes and restore AppImage notifications by @ksyasuda in #205
- fix(subtitles): recover canonical lines from ASS animation by @ksyasuda in #207
- fix(overlay): deduplicate secondary subtitle rendering by @ksyasuda in #208
- fix(launcher): restore Matroska thumbnails in Linux rofi picker by @ksyasuda in #210
- fix(character-dictionary): cache completed MeCab refreshes by @ksyasuda in #212
- fix(subtitles): improve secondary subtitle extraction and display by @ksyasuda in #215
- feat(release): track prerelease deltas and validate committed notes by @ksyasuda in #216
- fix(subtitles): recover positioned ASS word spacing and drop control debris by @ksyasuda in #217

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
