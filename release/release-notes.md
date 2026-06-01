## Highlights
### Fixed

- **Linux Overlay Stacking (XWayland / Wayland)**
  - The overlay no longer drops behind mpv on KDE Plasma and other non-Hyprland/Sway Wayland sessions; subtitle hover, pause-on-hover, and Yomitan lookups now work correctly on those desktops.
  - Stacking is now focus-sensitive: overlay stays managed while mpv is windowed, switches to non-interactive mode in fullscreen, and automatically yields to foreground windows (Settings, Yomitan, other apps) rather than covering them.
  - Startup glitches are resolved — no more display-sized overlay flash or black screen before playback begins.

- **Hyprland Overlay Placement (0.55+ / Lua configs)**
  - Overlay placement now works on Hyprland 0.55+ installations that use the new Lua config format; SubMiner detects Lua mode and uses the correct `hl.window_rule` dispatcher automatically.

- **macOS Overlay**
  - Fixed the subtitle overlay remaining click-through after pause-until-ready releases playback; hovering and Yomitan lookups resume normally.
  - Restored automatic mpv focus after closing Settings, AniList setup, and other modal windows so subtitles and playback keybinds work without clicking the player.

- **Manual Overlay Startup**
  - Starting the visible overlay manually from mpv now correctly attaches to playback, syncs the overlay window to mpv bounds on Linux/X11, and loads the current primary and secondary subtitles before revealing.

- **Playlist Transitions**
  - Advancing to the next mpv playlist item no longer triggers a second startup and tokenization delay; the overlay stays warm and visible subtitles are preserved across the transition.

- **Windows Launcher**
  - The `SubMiner mpv` shortcut on Windows now attaches the video to an already-running background app instead of spawning a duplicate warmup process.

- **Mouse Keybindings**
  - Side mouse buttons (`MBTN_BACK`, `MBTN_FORWARD`) and other mouse buttons can now be captured in the keybinding settings and work correctly at runtime.

### Docs

- **Troubleshooting Guides**
  - Hyprland overlay guide updated with both Lua (`hl.window_rule`) and legacy `hyprland.conf` window rule syntax, plus a note on automatic placement via `hyprctl`.
  - New KDE Plasma / Wayland section covering XWayland workarounds when launching mpv manually.
  - New Character Dictionary section covering name matching, inline portraits, and external-profile mode (no AniList login required).
  - Added a "See Also" index linking each feature to its own troubleshooting page.

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
