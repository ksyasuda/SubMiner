> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

<!-- prerelease-base-version: 0.19.0 -->

## Highlights
### Added

- **Sync Stats & History**
  - New **Sync Stats & History** window (tray menu) and `subminer sync <host>` command keep mining stats and watch history in sync across machines over SSH.
  - Syncing is safe to repeat: data merges without duplicates, and hosts with auto-sync enabled sync automatically in the background on a schedule, reporting results as overlay notifications.
  - Manual database snapshots (create, merge, reveal, delete) and connection testing cover one-off transfers, and Windows machines running the built-in OpenSSH Server work as sync remotes too; no setup beyond SSH access is needed on the remote side.

- **TsukiHime Subtitle Downloads**
  - Download Japanese and secondary-language subtitles for the current video directly from TsukiHime, mirroring the existing Jimaku flow: `Ctrl+Shift+T` opens an in-overlay search modal with separate tabs for the primary and secondary languages.
  - Matching releases are found automatically from the video filename; the chosen subtitle downloads and loads straight into mpv, no API key required.

- **Post-Playback History Menu**
  - After a watch-history episode ends or mpv closes, the fzf or rofi launcher returns to that series with options to play the previous episode, rewatch, play the next episode, select another episode, or quit SubMiner. Previous and Next continue across season directories.

### Changed

- **Clipboard-Video Shortcut**: The "append clipboard video to queue" shortcut is now configurable (`shortcuts.appendClipboardVideoToQueue`) instead of fixed.

### Fixed

- **Word Highlighting Accuracy**: Fixed several cases of incorrect word highlighting and annotations, including inconsistent part-of-speech exclusions on merged quote-particle tokens, missing annotations for rare kanji, katakana punctuation wrongly treated as non-kana noise, and certain kanji vocabulary being skipped for next-level ("N+1") highlighting.
- **Startup Playback Pausing Too Early**: Fixed playback resuming before subtitle processing had finished warming up, which could briefly show untranslated subtitles right after opening a video, most noticeable when resuming mid-episode.
- **Linux AppImage Crash Notification on Quit**: Fixed a spurious "Service Crash" desktop notification appearing after closing a video when running the Linux AppImage.
- **AnkiConnect Proxy Port Conflict**: Fixed video playback failing to start when another process already held the configured AnkiConnect proxy port; SubMiner now shows a notification explaining how to resolve the conflict instead of crashing.
- **Stats & Settings Reliability**: Hardened the stats server against malformed requests, stalled AniList lookups, media mismatches during word mining, and missing Yomitan connections; AnkiConnect settings validation now preserves valid custom configurations while safely falling back on invalid values instead of failing.

## What's Changed

- feat(shortcuts): make clipboard-video-append shortcut configurable by @ksyasuda in #158
- refactor(tokenizer): extract subtitle annotation filter into rule table by @ksyasuda in #162
- refactor(tsukihime): swap Animetosho backend for TsukiHime API by @ksyasuda in #165
- refactor: split anki-connect and stats-server resolvers into modules by @ksyasuda in #169
- feat(launcher): add post-playback history menu with previous episode by @ksyasuda in #170

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
