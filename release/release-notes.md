## Highlights
### Added
- **In-App Changelog**
  - View release notes without leaving the app, from the tray menu ("View Changelog") or the "What's New" button on update notifications.
  - Shows notes for the latest published release even when it's newer than your installed build, and falls back to the notes bundled with your install if the download fails.
  - Older versions fold automatically, your installed version is badged, and newer ones are tagged "New"; navigate with `J`/`K` or the arrow keys, `Enter` to expand/collapse, `R` to refresh, and `Esc` to close.

### Changed
- **Faster Subtitle Tokenization**
  - Subtitle lines are parsed and looked up roughly twice as efficiently, with results cached across lines so repeated words and grammar no longer re-query the dictionary.
  - Enabling a character dictionary no longer slows subtitle scanning as much, since name lookups now only check positions where a known name can actually start.
  - Fixed related accuracy issues along the way: readings that could go missing on certain word endings, subtitle text that stayed unannotated after mining a card, character names that could drop out of disambiguation rules, and halfwidth-katakana character names that weren't recognized or read correctly.

### Fixed
- **Large Character Dictionary Generation**
  - Big character dictionaries (long-running series like One Piece) no longer fail to install with a timeout error; the import time budget now scales with dictionary size instead of using a fixed 7-second limit.
  - The "Generating character dictionary" notification now shows real progress (character/page counts, image download progress with an ETA, name-processing progress) and an elapsed-time clock, so a long-running import no longer looks frozen.
- **Stats Deletion Responsiveness**
  - Deleting sessions, episodes, or library entries on the stats page no longer freezes the page or an active video player; deletes are now batched into a single transaction.
- **Subtitle Sidebar Clutter from Styled Subtitles**
  - Heavily typeset subtitles (karaoke openings/endings, stylized signs) no longer flood the subtitle sidebar with garbled vector-drawing text or duplicate "shadow" copies of the same line.
  - Subtitle text is now decoded consistently in one place, matching what mpv actually renders on screen, so it can no longer diverge or get cached inconsistently.
- **X11/XWayland Playback and Overlay Fixes**
  - Fixed a crash on the first fullscreen toggle when using an mpv `gpu-next` shader (e.g. ArtCNN) in X11/XWayland mode; SubMiner no longer forces mpv onto its older OpenGL renderer.
  - Fixed the overlay appearing oversized and offset from the video under fractional or mixed-monitor display scaling in X11/XWayland mode.

## What's Changed

- perf(tokenizer): single-pass Yomitan scan with cross-line caching and prefetch fixes by @ksyasuda in #185
- fix(subtitles): collapse duplicate ASS events and decode text once by @ksyasuda in #186
- feat(overlay): add in-app changelog modal by @ksyasuda in #187
- fix(playback): stop forcing legacy OpenGL renderer on X11 mpv backend by @ksyasuda in #188
- fix(dictionary): stop large character dictionaries from timing out by @ksyasuda in #189
- fix(overlay): handle X11 display scaling across monitors by @ksyasuda in #193
- fix(stats): batch deletes off the main thread by @ksyasuda in #194

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
