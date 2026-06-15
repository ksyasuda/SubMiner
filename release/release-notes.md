## Highlights
### Changed

- **Subtitle Delay Keybindings:** Overlay subtitle delay controls now match mpv's native bindings.
  - `z`, `Z`, and `x` adjust subtitle delay; `Ctrl+Shift+Left/Right` step to the adjacent subtitle and show the current delay on the OSD.
  - Removes the old SubMiner-specific adjacent-cue delay action in favor of mpv's built-in `sub-step`.

- **Update Notifications:** New installs now default to overlay-only update notifications instead of also sending a system notification.

### Fixed

- **Anki — Highlight Word:** The mined word is now correctly bolded in Kiku sentence and sentence-furigana fields even when the source Yomitan sentence did not already contain bold markup.

- **Anki — Lapis/Kiku Word Cards:** Word cards enriched through SubMiner now correctly include the word-and-sentence marker, restoring sentence context on the card front.

- **Anki — Windows:** Fixed two issues that surfaced after background app launches on Windows.
  - Audio clip and image generation now works correctly by recreating missing FFmpeg temp directories before export.
  - Known-word cache refreshes no longer fail when no deck is configured.

- **Desktop Notifications:** Restored the SubMiner app icon on system notifications that do not supply their own image.

- **Dictionary — Windows:** The character dictionary auto-sync on `SubMiner mpv` shortcut launches can now fall back to mpv's current video path when app media state is not yet ready.

- **Overlay — macOS Yomitan Popup:** Fixed focus and dismiss behavior for the Yomitan popup on macOS.
  - Popup focus is correctly restored after mining a card or reloading the popup.
  - Clicking on transparent overlay space now properly closes the popup and passes the click through to mpv, with no hide/reappear cycle.

- **Overlay — Linux Startup:** Fixed several edge cases that could leave the overlay unresponsive or drop subtitles at startup when auto-pause was active.
  - The overlay stays interactive during the initial render measurement gap.
  - Subtitles paint as plain text immediately on cache misses, before tokenization finishes.
  - Temporarily empty subtitle state is now re-parsed correctly before warm readiness resumes playback.

- **Overlay — Playlist Advance:** The visible overlay now stays interactive when mpv advances to the next playlist item, including when the next episode loads after the warm transition delay.

- **Overlay — Windows:** Fixed shaky subtitle-bar hover and click behavior when a video connects to an already-running background SubMiner instance.

- **Stats — AniList Search:** Manual AniList linking from the stats anime page now searches only the anime title, dropping any generated "Season N" suffix that was causing failed lookups.

- **Updates — Linux:** Improved Linux update reliability for managed support assets.
  - Updates now correctly install and refresh both the launcher runtime plugin copy and the rofi theme alongside AppImage and launcher updates.
  - Support-asset refreshes no longer touch unrelated SubMiner data directories, and plugin copies are staged safely before replacing the live runtime plugin.
  - Fresh installs now auto-install the managed runtime plugin and rofi theme from the bundled app on first launcher playback if either asset is missing.

### Docs

- **Linux Updates:** Documented how Linux update flows manage the launcher runtime plugin and rofi theme, and that the first launcher playback auto-installs any missing managed support assets from the bundled app.

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
