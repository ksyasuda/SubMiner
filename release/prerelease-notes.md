> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

<!-- prerelease-base-version: 0.19.4 -->

## Highlights
### Added

- Library Merge & Reassignment
  - Duplicate library entries for the same show can now be merged: pick entries in "Select" mode and use "Merge Selected" to combine sessions, mined cards, and watch time onto one card.
  - Episodes can be moved to a different library entry with a per-episode "→" button, fixing cases where a stray filename split off its own entry; manual assignments now survive later filename parsing, Jellyfin refreshes, and season repair.
  - Exact AniList matches with compatible seasons now merge automatically, and likely (fuzzy) matches surface as a dismissible "Possible duplicate" suggestion instead of merging silently.

- Duplicate Line Cleanup
  - The Vocabulary tab's new **Duplicates** button scans a chosen time window for the repeated-line bursts described under Fixed below and collapses each burst to a single line once you confirm it; a matching `subminer stats cleanup --duplicate-lines` command (with `--dry-run` and `--lookback-days <n>`) is available from the terminal.
  - Only the affected subtitle lines and the vocabulary counts they inflated are touched; watch time and lines-seen totals are left as recorded.

### Fixed

- Subtitle Duplication from Karaoke & Animated Signs
  - Typeset ASS karaoke and animated signs no longer flood the overlay, subtitle sidebar, immersion history, mined cards, or stats with repeated glyph fragments or per-frame duplicates; the complete authored line is recovered instead, without merging genuinely repeated dialogue or separately positioned signs.
  - The secondary overlay now shares the same deduplication logic as the primary overlay, so layered animation text no longer appears multiple times there or in what gets mined.
  - Vocabulary stats no longer count every animation frame of a karaoke opening as a separate line, which previously could push an OP lyric to the top of "Top Repeated Words."

- Anki Media Generation
  - Sentence-audio generation no longer times out on slow network-mounted video files with many subtitle and font streams, and a failed extraction now reports a clear error instead of a raw `ENOENT`.
  - Mined audio and animated AVIF clips now capture the subtitle line you actually mined, instead of whatever line happened to be on screen once slow audio extraction finished.

- Character Dictionary Performance & Notifications
  - Character dictionary generation, merged rebuilds, and imports no longer freeze the app on large dictionaries, and cached results are reused across launches instead of regenerating character data and portraits every time.
  - Desktop progress notifications, including on Linux AppImage installs, now update in place instead of flickering closed and reopening.

- Overlay Reliability
  - Overlay modals (settings, stats, etc.) now open promptly on the first shortcut press and appear above fullscreen mpv on macOS instead of switching Spaces or opening off-screen.
  - The overlay no longer gets stuck on "Overlay loading" indefinitely if mpv's connection stalls; it now retries and shows an actionable error after 30 seconds.
  - Fixed native Wayland drag-and-drop from file managers like Thunar, so subtitle and video files dropped on the overlay reach mpv.
  - Fixed system-wide mouse lag on Windows caused by the overlay's click-through handling and repeated mpv window lookups.

- Stats Dashboard
  - Deletes, library merges, video moves, and AniList reassignments no longer freeze the stats dashboard or rebuild lifetime totals from scratch; large deletes that used to take minutes now finish in milliseconds.
  - Vocabulary totals and charts now count all tracked vocabulary instead of just the first page, new-word history uses corrected daily rollups, calendar labels respect time zones west of UTC, and vocabulary cards refresh automatically after editing the word exclusion list.

- Linux Launcher Thumbnails
  - Fixed missing MKV thumbnails in the Linux rofi picker when the system thumbnailer only registers legacy Matroska MIME aliases.

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

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
