## Highlights
### Added
- **Anki Maturity Known-Word Highlighting**
  - Subtitle words you already know can now be color-coded by their Anki card maturity (new, learning, young, mature), like asbplayer's known-word coloring.
  - Enable it with `ankiConnect.knownWords.maturityEnabled` (or toggle it mid-session); tier colors and the "mature" day threshold are configurable, and the in-session help legend shows the active colors.
- **Cross-Machine Sync for Stats & Watch History**
  - Sync immersion stats and watch history between machines over SSH from a new Sync window (tray menu → Sync Stats & History, or `subminer sync --ui`) or the CLI (`subminer sync <host>`).
  - Save multiple devices with per-host sync direction, run one-click syncs with live progress, and take manual database snapshots for backup or transfer.
  - Windows remotes are supported over OpenSSH, and hosts can auto-sync in the background on a schedule, even during playback.
- **History Menu After Playback**
  - After a watch-history episode ends (or mpv closes), the fzf/rofi launcher now offers to play the previous or next episode, rewatch, pick another episode, or quit, right from where you left off. Previous/Next continue across season folders.
- **Delete Entire Library Titles from Stats**
  - The stats Library detail view now has a "Delete Entry" action that removes a whole title in one step, episodes, sessions, subtitle lines, rollups, cover art, and vocabulary counts, instead of clearing it episode by episode.
  - Delete progress (sessions, episodes, or whole titles) now shows app-wide with a progress bar and status toast visible from any tab or window.
- **TsukiHime English Subtitle Downloads**
  - Download subtitles for the currently playing video directly from TsukiHime, with Japanese loaded as the primary track and your configured secondary language alongside it.

### Changed
- **Configurable Clipboard-Video Shortcut**
  - The "append clipboard video to queue" shortcut is now configurable via `shortcuts.appendClipboardVideoToQueue` instead of fixed.

### Fixed
- **AniList Season Matching**
  - Season 2+ episodes now resolve to the correct AniList entry instead of silently falling back to season 1. SubMiner follows AniList's sequel relations to find the right season, and cover art and watch progress now use the same season-aware match.
  - If a season still can't be found, SubMiner no longer force-writes progress or a cover to the season 1 entry, it skips the update and points you to a manual AniList override, which now fixes the character dictionary and watch progress together.
  - Manual overrides now stay applied consistently across every episode in a season folder, even when filenames guess differently episode to episode.
- **Subtitle Highlighting Accuracy**
  - Fixed several known-word/annotation edge cases: part-of-speech exclusions now apply consistently to merged quote-particle tokens, annotations for rarer kanji are preserved, katakana punctuation is no longer mistaken for plain kana, and a specific noun-tagging case no longer loses its known+1 highlight.
- **AnkiConnect Proxy Port Conflicts**
  - Fixed a crash on video startup when another process already held the configured AnkiConnect proxy port; you'll now get a notification explaining how to resolve it instead.
- **AppImage Crash Notification on Quit**
  - Fixed a spurious "Service Crash" desktop notification appearing after closing a video when running the Linux AppImage.
- **Startup Playback Pause Timing**
  - Fixed playback occasionally resuming a couple seconds before subtitle tokenization actually finished warming up, most noticeable when resuming mid-episode or when a subtitle appears in the first two seconds.
- **Stats Library Cover After Relinking**
  - Fixed the stats Library grid showing a stale cover image after relinking a title to a different AniList entry.
- **Faster Stats Deletes and Vocabulary Tab**
  - Deleting sessions, episodes, and titles from stats is now dramatically faster and no longer stalls playback while it runs; the Vocabulary tab also loads much faster.
  - The first launch after updating runs a one-time database migration (a few seconds, database grows about 20%); no action needed.
- **Settings Validation and Stats Server Hardening**
  - Invalid AnkiConnect settings now fall back safely with a warning instead of silently breaking, and the stats server is hardened against malformed requests, stalled AniList searches, and other edge cases that could previously crash it.
- **Rofi Prompt Spacing**
  - Fixed rofi menu prompts running into the search placeholder text with no space between them.

## What's Changed

- feat(shortcuts): make clipboard-video-append shortcut configurable by @ksyasuda in #158
- refactor(tokenizer): extract subtitle annotation filter into rule table by @ksyasuda in #162
- refactor(tsukihime): swap Animetosho backend for TsukiHime API by @ksyasuda in #165
- refactor: split anki-connect and stats-server resolvers into modules by @ksyasuda in #169
- feat(launcher): add post-playback history menu with previous episode by @ksyasuda in #170
- Anki maturity-based known-word highlighting by @ksyasuda in #172
- fix(anilist): resolve later seasons via sequel relations, not title guessing by @ksyasuda in #173
- feat(stats): add library entry deletion and app-wide delete progress by @ksyasuda in #174

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
