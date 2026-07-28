> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

<!-- prerelease-base-version: 0.19.0 -->

## Highlights
### Added

- **Sync Stats & History**
  - New **Sync Stats & History** window (tray menu) and `subminer sync <host>` command keep mining stats and watch history in sync between machines over SSH, with saved devices, per-host sync direction, and live stage-by-stage progress.
  - Merges are safe to repeat: data combines without duplicates, and hosts with auto-sync enabled sync automatically in the background on a schedule, reporting results as overlay notifications.
  - Manual snapshot tools (create, merge, reveal, delete) and connection testing cover one-off transfers; Windows machines running the built-in OpenSSH Server work as sync remotes too, with no setup needed beyond SSH access. Power users can script transfers directly with `--push`/`--pull`, `--check`, `--snapshot`/`--merge`, and `--json` flags.

- **TsukiHime Subtitle Downloads**
  - Download Japanese and secondary-language subtitles for the current video directly from TsukiHime, mirroring the existing Jimaku flow.
  - Press `Ctrl+Shift+T` to search by tabs for the primary and secondary languages; the matching release is found automatically from the video filename and loads straight into mpv, no API key required.

- **Post-Playback History Menu**
  - After a watch-history episode ends or mpv closes, the fzf/rofi launcher returns to that series with options to play the previous or next episode, rewatch, pick another episode, or quit SubMiner.
  - Previous/Next continue across season directories, so you can binge a show without manually browsing folders.
  - The menu shown right after picking a series from `subminer -H` now offers the previous episode too, matching the post-playback menu.

- **Known-Word Highlighting by Anki Maturity**
  - Subtitle highlights for known words can now be colored by Anki card maturity (new, learning, young, mature), similar to asbplayer. Enable it with `ankiConnect.knownWords.maturityEnabled`, or toggle it live during a session.
  - The mature-interval threshold and the four tier colors are configurable, and the in-session help legend shows the active tier colors while maturity highlighting is on.
  - Tiers follow Anki's own card state: a lapsed card correctly shows as learning rather than young, and a note is treated as mature if any of its cards are mature. Stats and other known-word tools stay accurate with this new data.

### Changed

- **Clipboard-Video Shortcut**
  - The "append clipboard video to queue" shortcut is now configurable via `shortcuts.appendClipboardVideoToQueue` instead of being fixed.

### Fixed

- **Word Highlighting Accuracy**
  - Fixed several incorrect word highlighting and annotation cases: inconsistent part-of-speech exclusions on merged quote-particle tokens, missing annotations for rare kanji, katakana punctuation wrongly treated as non-kana noise, and certain kanji vocabulary skipped for next-level ("N+1") highlighting.

- **Character Dictionary Season Overrides**
  - Manual AniList overrides for a series now stay in effect for every episode in the same season folder, even when individual episode filenames produce different automatic guesses.

- **Startup Playback Pausing Too Early**
  - Fixed playback resuming before subtitle processing finished warming up, which could briefly show untranslated subtitles right after opening a video.
  - Most noticeable when resuming mid-episode or when a subtitle cue starts within the first couple of seconds.

- **Linux AppImage Crash Notification on Quit**
  - Fixed a spurious "Service Crash" desktop notification appearing after closing a video when running the Linux AppImage.
  - If needed, the mount-keepalive behavior behind this fix can be disabled with `SUBMINER_NO_APPIMAGE_MOUNT_KEEPALIVE=1`.

- **AnkiConnect Proxy Port Conflict**
  - Fixed video playback failing to start when another process already held the configured AnkiConnect proxy port; SubMiner now shows a notification explaining how to resolve the conflict instead of crashing.

- **Stats & Settings Reliability**
  - Fixed session stats reporting zero known words after the known-word cache gained maturity tiers.
  - Hardened the stats server against malformed requests, stalled AniList lookups, media mismatches during word mining, and missing Yomitan connections.
  - AnkiConnect settings validation now preserves valid custom configurations while safely falling back on invalid values instead of failing.

- **Rofi Menu Prompt Spacing**
  - Rofi menu prompts now keep a space between the prompt label and the input field instead of crowding the search placeholder text.

## What's Changed

- feat(shortcuts): make clipboard-video-append shortcut configurable by @ksyasuda in #158
- refactor(tokenizer): extract subtitle annotation filter into rule table by @ksyasuda in #162
- refactor(tsukihime): swap Animetosho backend for TsukiHime API by @ksyasuda in #165
- refactor: split anki-connect and stats-server resolvers into modules by @ksyasuda in #169
- feat(launcher): add post-playback history menu with previous episode by @ksyasuda in #170
- Anki maturity-based known-word highlighting by @ksyasuda in #172

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
