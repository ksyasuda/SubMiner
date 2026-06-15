> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

<!-- prerelease-base-version: 0.17.0 -->

## Highlights
### Changed

- **Subtitle Delay Shortcuts:** Overlay subtitle delay controls now match mpv's native defaults.
  - `z`, `Z`, and `x` adjust `sub-delay`; `Ctrl+Shift+Left/Right` run native `sub-step` and show the current delay on the OSD.
  - The previous SubMiner-only adjacent-cue delay action has been removed.

- **Update Notifications:** New installs now default to overlay-only update notifications instead of overlay plus system notifications.

### Fixed

- **Anki Card Enrichment:** Fixed two issues where card fields were not populated correctly after mining.
  - Highlight Word now bolds the mined word in Kiku sentence and sentence-furigana fields even when the source Yomitan sentence has no existing bold markup.
  - Lapis and Kiku word cards enriched through SubMiner now include the word-and-sentence marker, restoring sentence context on the card front.

- **Windows Overlay:** Fixed shaky hover and click behavior on the subtitle bar when a video attaches to an already-running SubMiner instance.

- **Windows Anki & Media:** Fixed two issues affecting Windows users running SubMiner in background-launch mode.
  - Known-word cache refreshes no longer fail when no deck is configured.
  - Audio and image clipping now works correctly by recreating missing FFmpeg temp directories before processing.

- **Windows Character Dictionary:** The character dictionary auto-sync now correctly falls back to mpv's current video path on Windows when app media state is not yet ready.

- **Linux Support Assets:** Linux updates now create and refresh both managed support assets: the launcher runtime plugin copy and the rofi theme.
  - First playback on a fresh Linux install auto-installs those bundled assets before mpv starts if either one is missing.
  - Asset refreshes leave unrelated SubMiner data directories untouched and stage plugin copies before replacing the live runtime plugin.

- **Linux Visible Overlay Startup:** Auto-paused visible overlay startup stays fully interactive during the first measurement gap.
  - Startup subtitle cache misses paint raw text before tokenization finishes, and temporarily empty mpv subtitle reads refresh parsed cues before warm readiness resumes playback.

- **Playlist Transitions:** The visible overlay stays active while mpv advances to the next playlist item, including when the next episode loads after the warm transition delay.

- **macOS Yomitan Popup Focus:** Yomitan popup focus is restored after card mining or popup reload.
  - Clicking transparent overlay space now closes the popup and returns passthrough to mpv without a hide/reappear cycle.

- **Stats AniList Search:** Manual AniList linking from the stats page now strips generated `Season N` suffixes before searching, so the base anime title is used.

- **Desktop Notifications:** System notifications now show the SubMiner app icon when no custom notification image is provided.

- **Release Notes:** GitHub release `What's Changed` and `New Contributors` attribution sections are preserved when CI regenerates release notes from committed changelog output.

### Docs

- **Linux Update Flow:** Documented that Linux update flows manage the launcher runtime plugin copy and rofi theme from `subminer-assets.tar.gz`, and that normal playback auto-installs those managed support assets if either one is missing.

## What's Changed

- Replace subtitle delay actions with native mpv keybindings by @ksyasuda in #120
- fix(stats): strip Season N suffix from AniList title searches by @ksyasuda in #121
- fix(overlay): preserve visible state across playlist item transitions by @ksyasuda in #124
- fix(overlay): restore macOS Yomitan popup focus without breaking click-away by @ksyasuda in #125
- fix(linux): auto-install managed plugin copy; include in asset updates by @ksyasuda in #127
- Fix Windows Anki startup and overlay regressions by @ksyasuda in #128

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
