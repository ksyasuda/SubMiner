> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

## Highlights
### Added

**Auto-Updater:** SubMiner can now check for and apply updates from the system tray or by running `subminer -u`. Checks include checksum verification, configurable notifications, and an opt-in channel for prerelease builds. The `subminer` launcher and Linux rofi theme are also updated automatically.

**First-Run Setup:** A new optional setup flow installs Bun and the `subminer` command-line launcher on Linux, macOS, and Windows. Windows users get a `subminer.cmd` PATH shim so `subminer` works in any terminal without manually adding `SubMiner.exe` to PATH.

### Fixed

**macOS Overlay:** Significantly improved overlay focus and stability — the overlay now hides when mpv loses focus or is minimized, stays stable through transient window-tracking misses, remains correctly layered during stats mouse passthrough, and opens over fullscreen mpv without switching Spaces. Passthrough is also fixed so mpv controls stay clickable before hovering a subtitle bar. Background tracking overhead is reduced while mpv is stably focused.

**Subtitle Sync Modal:** Fixed a macOS issue where opening the subtitle sync modal would flash and disappear on the first attempt, or leave stale state after syncing.

**Controller:** Controller config and debug shortcuts now stay closed while controller support is disabled, with a notice to enable `controller.enabled`. Learn mode can be entered from the edit pencil or binding badge, remaps are saved per controller profile, and individual bindings can be reset to their defaults.

**AniList Progress:** Progress threshold checks now use fresh playback position data, so updates fire correctly when playback reaches or skips past the watched threshold. Season-specific results are preferred for multi-season files, and a clear message is shown when the matched season is not in Planning or Watching status.

**Character Dictionary:** Cached media matches are reused when loading a title with an existing snapshot, avoiding redundant AniList search requests on repeat visits.

**Updater — Linux:** The tray app now uses GitHub release metadata for update checks instead of the native Electron updater, preventing crashes. `subminer -u` performs updates independently of any running tray instance and correctly reports "up to date" without downloading assets when no newer release exists.

**Updater — macOS:** Update dialogs now come to the front when launched from `subminer --update`. Builds that cannot install native updates show a manual-install message instead of an inapplicable restart prompt. Signed macOS builds remain on the native updater path without triggering premature Squirrel install checks.

**Setup — macOS:** First-run setup now recognizes existing `subminer` launcher installs in Homebrew or user PATH directories, and manual setup avoids writing into Homebrew-owned paths. `subminer app --setup` opens the setup flow even when SubMiner is already running in the background. The standalone setup app quits after completing first-run setup, returning control to the terminal.

**Launcher — Linux:** First-run launcher installs are now built with a valid Bun shebang, fixing installs that previously failed silently.

**Tray App:** Fixed several lifecycle issues with tray-launched Yomitan settings: the tray stays running when settings are closed; settings loading no longer blocks other tray actions; the settings window uses a close-only menu to prevent accidentally quitting the tray app; an in-page close button is provided on Hyprland where native window controls are unavailable; the embedded popup preview is disabled to prevent renderer hangs during sidebar navigation; extension refreshes at startup are serialized to prevent race conditions; and the session help modal can now close correctly without mpv running.

**Build — Linux Install:** Fixed one-shot `make clean build install` flows so the install step correctly picks up the AppImage produced earlier in the same make invocation.

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
