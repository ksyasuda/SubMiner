> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

<!-- prerelease-base-version: 0.17.0 -->

## Highlights
### Changed

- **Subtitle Delay Shortcuts:** Overlay subtitle delay controls now match mpv's native defaults.
  - `z`, `Z`, and `x` adjust `sub-delay`; `Ctrl+Shift+Left/Right` run native `sub-step` and show the current delay on the OSD.
  - The previous SubMiner-only adjacent-cue delay action has been removed.

- **Update Notifications:** New installs now default update notifications to overlay-only instead of overlay plus system notifications.

### Fixed

- **Linux Support Assets:** Linux updates now create and refresh both managed support assets: the launcher runtime plugin copy and the rofi theme.
  - First launcher playback on fresh Linux installs auto-installs those bundled support assets before mpv starts when either asset is missing.
  - Support-asset refreshes leave unrelated SubMiner data directories alone and stage plugin copies before replacing the live runtime plugin.

- **Linux Visible Overlay Startup:** Auto-paused visible overlay startup stays interactive during the first measurement gap.
  - Startup subtitle cache misses paint raw text before tokenization finishes, and temporarily empty mpv subtitle reads refresh parsed cues before synthetic warm readiness resumes playback.

- **Playlist Transitions:** The visible overlay stays active while mpv advances to the next playlist item, including when the next episode loads after the warm transition delay.

- **macOS Yomitan Popup Focus:** Yomitan popup focus is restored after card mining or popup reload.
  - Clicking transparent overlay space now closes the popup, then returns passthrough to mpv without a hide/reappear cycle.

- **Stats AniList Search:** Manual AniList linking from the stats anime page now drops generated `Season N` suffixes before searching, so automatic searches use the base anime title.

- **Desktop Notifications:** System notifications now show the SubMiner app icon when no custom notification image is provided.

- **Release Notes:** GitHub release `What's Changed` and `New Contributors` attribution sections are preserved when CI regenerates release notes from committed changelog output.

### Docs

- **Linux Update Flow:** Documented that Linux update flows manage the launcher runtime plugin copy and rofi theme from `subminer-assets.tar.gz`, and that normal launcher playback auto-installs those managed support assets if either one is missing.

## What's Changed

- Replace subtitle delay actions with native mpv keybindings by @ksyasuda in #120
- fix(stats): strip Season N suffix from AniList title searches by @ksyasuda in #121
- fix(overlay): preserve visible state across playlist item transitions by @ksyasuda in #124
- fix(overlay): restore macOS Yomitan popup focus without breaking click-away by @ksyasuda in #125
- fix(linux): auto-install managed plugin copy; include in asset updates by @ksyasuda in #127

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
