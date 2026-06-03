## Highlights
### Changed

- **Yomitan**: Updated the bundled Yomitan to the latest revision.
  - Picks up the newest lookup improvements and fixes from the SubMiner Yomitan fork.

### Fixed

- **Anki / Animated AVIF**: Clips with word audio no longer start or end early.
  - Clip boundaries are now snapped to the nearest AVIF frame edge, keeping audio lead-in and playback in sync.

- **macOS Overlay**: Resolved several interactivity and focus issues triggered by autoplay and modal windows.
  - After autoplay starts with "wait for overlay to be ready" enabled, subtitles are immediately hoverable and Yomitan lookups work - no longer require an extra click to activate.
  - After any modal closes (Settings, Stats, sidebar, etc.), the overlay and subtitles reappear automatically and mpv keyboard shortcuts (pause, seek, etc.) are restored to mpv right away, including in native fullscreen.

- **Hyprland Fullscreen Overlay**: Fixed overlay alignment when mpv is fullscreen on Hyprland.
  - Compositor client bounds are now verified before positioning, so the stats panel, modals, and subtitle sidebar no longer shift below the mpv window.

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
