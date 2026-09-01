## Highlights
### Fixed

- **Anki Card Update Progress**: The update spinner now stays visible until audio and image updates actually finish, so you won't mistake an in-progress update for a failure.
- **Word-Card Field Enrichment**: Word-card enrichment now reliably writes sentence text and audio into whichever AnkiConnect fields you've configured, while the dedicated Lapis/Kiku sentence-card and audio-card actions still use their expected field names.
- **Overlapping Subtitles**:
  - Lines that start while another is still on screen now show together instead of staying hidden until you switch tracks or seek.
  - Subtitles shown at the same time now stack by their authored screen position, with signs and song lyrics above dialogue.
  - Half-size ASS furigana no longer shows up as if it were its own subtitle line.
- **YouTube Auto-Generated Captions**:
  - Captions now follow their intended timing instead of drifting off sync.
  - Long speech is paged across two rows instead of piling into a wall of text.
  - Timed sound cues like `[音楽]` no longer linger over later dialogue.

## What's Changed

- fix(anki): keep overlay progress visible through card updates by @ksyasuda in #218
- fix(youtube): keep auto captions on screen for their full span by @ksyasuda in #219
- fix(subtitles): keep overlapping lines that join an already active cue by @ksyasuda in #221
- fix(anki): respect configured fields for word-card enrichment by @ksyasuda in #223

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
