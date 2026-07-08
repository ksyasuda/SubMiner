> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

<!-- prerelease-base-version: 0.18.0 -->

## Highlights
### Added
- **Watch History Browser**
  - New `subminer -H` / `--history` command to browse your local watch history, replay the last episode, jump to the next one, or pick an episode via fzf/rofi.
  - The rofi picker now shows AniList cover art for each show, making it easier to spot the right title at a glance.
- **Card Audio Normalization**
  - Audio extracted for Anki cards is now volume-normalized by default for more consistent playback loudness.
  - If you prefer the original source volume, disable it via the new `ankiConnect.media.normalizeAudio` setting.

### Changed
- **New App Icon**
  - SubMiner now ships pixel-art submarine artwork, contributed by an anonymous community member.
  - Applied across the app icon, tray icon, notifications, README, docs site, and stats page.

### Fixed
- **Character Name Highlighting in Subtitles**
  - Fixed unspaced Japanese names (e.g. 東紫乃, 渡辺真奈美) being split at the wrong point, which left surnames like 東 and 渡辺 without their character portrait or hover lookup.
  - Fixed names getting cut off or losing their highlight when caught by the subtitle scanner's punctuation handling or by conflicting grammar tagging.
  - No action needed — existing data upgrades automatically the next time a matching name is seen.
- **Known-Word Highlighting**
  - Words are no longer marked "known" (green) just because they share spelling with a known Anki card that actually teaches a different reading (e.g. 床 read as とこ no longer falsely matches a known 床/ゆか card).
- **Unparsed Subtitle Text**
  - Subtitle text the dictionary can't recognize (like a truncated verb form) is now still hoverable for lookup and correctly counted toward a sentence's difficulty, instead of showing as dead, non-interactive text.
- **Kiku Manual Field Grouping**
  - Fixed the field-grouping dialog getting stuck invisible behind fullscreen video on Hyprland/Wayland, and failing silently on repeated attempts after the first use.
  - Fixed a duplicate "Field grouping cancelled" notification appearing when grouping was cancelled via the trigger shortcut.
- **Secondary Subtitles**
  - Karaoke-style secondary subtitles (common in opening/ending songs) no longer spam dozens of lines down the screen; repeated lines are now collapsed and the subtitle area is capped to a strip at the top.
- **YouTube Extraction**
  - Fixed direct YouTube stream extraction occasionally corrupting the stream URL and causing failed audio/video capture.
- **Background Stats Server**
  - Launching SubMiner in the background now correctly auto-starts the stats server when enabled, and won't start a duplicate if one's already running.
- **Stats Trend Charts**
  - All trend chart titles now show by default, with the ability to hide specific titles (remembered across sessions) and cap how many top titles a chart displays.

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

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
