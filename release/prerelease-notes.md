> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

<!-- prerelease-base-version: 0.18.0 -->

## Highlights
### Added
- **Watch History Browser**
  - New `subminer -H` / `--history` command lets you browse your local watch history, replay the last episode, jump to the next one, or pick an episode via fzf or rofi.
  - The rofi picker now shows AniList cover art for each show, making it easier to spot the right title at a glance.
- **Card Audio Normalization**
  - Audio extracted for Anki cards is now volume-normalized by default, giving more consistent playback loudness across cards.
  - Prefer the original source volume? Disable it via the new `ankiConnect.media.normalizeAudio` setting.

### Changed
- **New App Icon**
  - SubMiner now ships pixel-art submarine artwork contributed by an anonymous community member.
  - Applied across the app icon, tray icon, notifications, README, docs site, and stats page.
- **Launcher Preview Layout**
  - fzf previews in the launcher now sit below the menu instead of beside it, giving long titles and metadata more horizontal room.

### Fixed
- **Character Name Highlighting in Subtitles**
  - Fixed unspaced Japanese names (e.g. 東紫乃, 渡辺真奈美) being split at the wrong point, which left surnames like 東 and 渡辺 without their character portrait or hover lookup.
  - Fixed names getting cut off or losing their highlight when caught by the subtitle scanner's punctuation handling, misclassified by grammar tagging, or swallowed entirely by a longer generic dictionary match (e.g. ヨータ disappearing inside a false とヨー match).
  - Fixed a single unrecognized word in a subtitle line (like a stray interjection) causing character-name highlighting to drop for the whole line instead of just that word.
  - No action needed — existing data upgrades automatically the next time a matching name is seen.
- **Known-Word Highlighting**
  - Words are no longer marked "known" (green) just because they share spelling with a known Anki card that actually teaches a different reading (e.g. 床 read as とこ no longer falsely matches a known 床/ゆか card).
  - Kanji words are also no longer marked known just because a different mined word happens to share their reading (e.g. 渓谷/けいこく no longer falsely matches a known 警告/けいこく card).
  - Single-kana grammar tokens (particles like よ, え) no longer borrow an unrelated card's reading and get falsely painted as known.
  - Stats sessions now correctly reflect known-word counts again after the reading-aware matching upgrade, instead of showing 0 everywhere.
- **Annotation Highlighting Refinements**
  - Restored frequency/JLPT highlighting and vocabulary-stat counting for words like 確かに and やはり, which were wrongly treated as grammar noise.
  - Kanji nouns that MeCab tags as "non-independent" (e.g. 日, 点, 以外) also keep their highlighting and stats counting again.
  - Suffix-only tokens (e.g. さん, れる) are now excluded from JLPT/frequency highlighting by default to match how particles and interjections are treated; known-word highlighting for them still works, and this is configurable.
- **Unparsed Subtitle Text**
  - Subtitle text the dictionary can't recognize (like a truncated verb form) is now still hoverable for lookup and correctly counted toward a sentence's difficulty, instead of showing as dead, non-interactive text.
- **Kiku Manual Field Grouping**
  - Fixed the field-grouping dialog getting stuck invisible behind fullscreen video on Hyprland/Wayland, and failing silently on repeated attempts after the first use.
  - Fixed a timed-out or failed grouping request leaving an invisible, stuck dialog covering the video; it now closes automatically so the overlay recovers.
  - Fixed a duplicate "Field grouping cancelled" notification appearing when grouping was cancelled via the trigger shortcut, and added a proper error message for the previously-silent case where the original card can no longer be loaded.
- **Secondary Subtitles**
  - Karaoke-style secondary subtitles (common in opening/ending songs) no longer spam dozens of lines down the screen; repeated lines are now collapsed and the subtitle area is capped to a strip at the top.
- **YouTube Extraction**
  - Fixed direct YouTube stream extraction occasionally corrupting the stream URL and causing failed audio/video capture.
- **Background Stats Server**
  - Launching SubMiner in the background now correctly auto-starts the stats server when enabled, and won't start a duplicate if one's already running.
- **Stats Trend Charts**
  - All trend chart titles now show by default, with the ability to hide specific titles (remembered across sessions) and cap how many top titles a chart displays.
- **Stats Cover Art**
  - Cover art now loads as soon as a series starts playing instead of waiting for your first visit to its detail page, so the stats timeline shows artwork right away.
  - Existing series missing art are backfilled automatically the next time you open the stats page.

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
- fix(overlay): keep frequency/JLPT highlight for kanji non-independent nouns by @ksyasuda in #150
- fix(tokenizer): greedy name pre-pass to prevent generic matches swallowing character names by @ksyasuda in #151

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
