> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

<!-- prerelease-base-version: 0.19.0 -->

## Highlights
### Added

- **Sync Stats & History**
  - Keep mining stats and watch history in sync across machines over SSH, from a new **Sync Stats & History** window (tray menu) or the `subminer sync <host>` command.
  - Syncing merges data safely, so nothing is duplicated even if you sync the same machines repeatedly, and hosts with auto-sync enabled sync in the background on a schedule, reporting results as overlay notifications.
  - Manual database snapshots (create, merge, reveal, delete) cover one-off transfers, and Windows machines running the built-in OpenSSH Server can be used as sync remotes too. No setup beyond SSH access is required on the remote side.

- **TsukiHime Subtitle Downloads**
  - Download Japanese and secondary-language subtitles for the current video directly from TsukiHime, mirroring the existing Jimaku flow: `Ctrl+Shift+T` opens an in-overlay search modal with separate tabs for the primary and secondary languages.
  - Matching releases are found automatically from the video filename; the chosen subtitle downloads and loads straight into mpv, no API key required.

### Changed

- **Clipboard-Video Shortcut**: The "append clipboard video to queue" shortcut is now configurable (`shortcuts.appendClipboardVideoToQueue`) instead of fixed.

### Fixed

- **Word Highlighting Accuracy**: Fixed several cases of incorrect word highlighting and annotations, including inconsistent part-of-speech exclusions on merged quote-particle tokens, missing annotations for rare kanji, katakana punctuation wrongly treated as non-kana noise, and certain kanji vocabulary being skipped for next-level ("N+1") highlighting.
- **Startup Playback Pausing Too Early**: Fixed playback resuming before subtitle processing had finished warming up, which could briefly show untranslated subtitles right after opening a video, most noticeable when resuming mid-episode.
- **Linux AppImage Crash Notification on Quit**: Fixed a spurious "Service Crash" desktop notification appearing after closing a video when running the Linux AppImage.

## What's Changed

- feat(shortcuts): make clipboard-video-append shortcut configurable by @ksyasuda in #158
- refactor(tokenizer): extract subtitle annotation filter into rule table by @ksyasuda in #162
- refactor(tsukihime): swap Animetosho backend for TsukiHime API by @ksyasuda in #165

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
