## Highlights
### Added

- **Sentence Audio Normalization**
  - Generated sentence audio is now normalized to -23 LUFS by default, giving mined clips consistent volume across shows.
  - Clips captured from playback can also mirror mpv's software volume curve, with a limiter to prevent clipping when boosted.
  - Both behaviors are controlled independently and can be turned off in the Anki Connect media settings.

- **Watch History Browser**
  - Added `subminer -H` / `--history` to browse local watch history, replay the last episode, continue to the next one, or jump to any past episode.
  - Works with fzf or rofi; the rofi picker shows AniList cover art already stored in the stats database.

### Changed

- **Known-Word Highlighting Accuracy**
  - Highlighting now compares subtitle and Anki-card readings, so it no longer confuses homographs like 床/とこ vs 床/ゆか or unrelated kanji that happen to share a reading.
  - Standalone suffix words such as さん or れる are now excluded from JLPT/frequency/N+1 annotations by default, matching how particles and interjections are already treated (configurable if you'd rather keep them annotated).
  - Cards without readings still fall back to word-only matching; the highlighting cache upgrades automatically with no action needed.

- **Stats Trend Charts**
  - Overhauled trend charts with persisted title visibility, per-chart title limits, "top" and "most recent" ranking modes, an option to show or hide empty days, calendar-aligned periods, and sortable multi-column tooltips.

- **New App Icon**
  - Replaced the SubMiner icon with new pixel-art submarine artwork contributed by an anonymous community member, now used across the app icon, tray, notifications, README, docs site, and stats page.

- **Launcher Preview Layout**
  - fzf previews now sit below the launcher menu instead of beside it, giving long titles and metadata more horizontal room.

### Fixed

- **Character Name Recognition**
  - Character dictionaries now correctly split unspaced AniList native names and validate readings, so overlay portraits, highlights, and hover lookups work reliably even without MeCab installed.
  - Name matches survive punctuation and unmatched text and no longer get overridden by generic dictionary matches or wrongly split from longer words like 空気.
  - Existing installs regenerate automatically and upgrade to exact splits once MeCab is available — no action needed.

- **Frequency & JLPT Highlighting Coverage**
  - Content adverbs (確かに, やはり) and kanji nouns MeCab tags as non-independent (日, 点, 以外) are now correctly included in frequency/JLPT highlighting and vocabulary stats.
  - Lexicalized kana expressions like かといって keep their frequency annotations, while interjections, pronouns, and pure grammar fragments are still filtered out as noise.

- **Unparsed Text Hover Lookup**
  - Subtitle text Yomitan can't fully parse — truncated inflections like とこ戻ろ… or elongation runs like ぅ～ — is hoverable again for dictionary lookup.
  - These runs stay excluded from frequency/JLPT highlighting and vocabulary stats, same as bracketed captions and punctuation-only text.

- **Karaoke-Style Secondary Subtitles**
  - Secondary subtitles no longer flood the screen with dozens of one-syllable lines during karaoke-style openings and endings.
  - Repeated events now collapse into a single line, and the secondary subtitle area stays capped to a strip at the top.

- **Kiku Field Grouping Reliability**
  - The manual field-grouping dialog now stays above fullscreen mpv on Hyprland/Wayland and keeps working across repeated attempts.
  - Abandoned grouping windows close automatically after a timeout or failure, and each attempt now reports a clear success or error, including when the original card can no longer be loaded.

- **Background Stats Server Startup**
  - Background `subminer app` launches now start the stats server automatically when enabled, and skip startup if one is already running.

- **AniList Cover Art Timing**
  - Stats now fetches the best-match AniList cover as soon as a new series starts playing, so artwork appears in the timeline immediately instead of only after visiting the series page.
  - Existing series missing art get backfilled automatically on the next Stats page visit.

- **YouTube Direct Stream Playback**
  - Fixed direct YouTube stream extraction so mpv's EDL stream URLs are parsed correctly, preventing corrupted signed video URLs and the resulting ffmpeg 403 errors.

## What's Changed

- fix(youtube): parse mpv EDL stream URLs with byte-length guards by @ksyasuda in #134
- Normalize generated Anki audio by default by @ksyasuda in #135
- feat(launcher): add -H/--history command to browse local watch history by @ksyasuda in #136
- fix(overlay): prevent field grouping modal from freezing overlay on Hyprland by @ksyasuda in #138
- fix(overlay): collapse karaoke syllable spam in secondary subtitles by @ksyasuda in #139
- feat(stats): Trends dashboard overhaul — title visibility, ranking modes, calendar-accurate windows, tooltips by @ksyasuda in #140
- feat(branding): replace app icon with contributed pixel-art set by @ksyasuda in #141
- feat(anki): reading-aware known-word matching (cache v3) by @ksyasuda in #142
- fix(stats): start stats server on background app launch by @ksyasuda in #144
- fix(tokenizer): keep unparsed Yomitan tokens hoverable by @ksyasuda in #145
- fix(overlay): resolve unspaced Japanese name splits and scan recovery by @ksyasuda in #146
- fix(tokenizer): prevent grammar tokens from borrowing known-word highlight via unrelated readings by @ksyasuda in #147
- fix(stats): fetch cover art eagerly at session start instead of on series page visit by @ksyasuda in #148

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
