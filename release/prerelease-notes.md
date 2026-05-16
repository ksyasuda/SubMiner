> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

## Highlights
### Added

- **Auto-Update:** Tray and `subminer -u` command-line update checks for new SubMiner releases, with app and launcher update prompts, checksum verification, configurable update notifications, and an opt-in prerelease channel for beta and RC builds.
- **First-Run Setup:** Guided setup flow to install Bun and the `subminer` command-line launcher on Linux, macOS, and Windows. On Windows, a `subminer.cmd` PATH shim is installed so you can type `subminer` in any terminal without adding the main executable to PATH.

### Fixed

- **macOS Overlay:** Transient mpv window appearances no longer incorrectly hide the subtitle overlay; minimizing mpv still hides it as expected. mpv controls are also now clickable before hovering a subtitle bar.
- **Subtitle Sync Modal:** Opening the subtitle sync panel on macOS no longer flashes and dismisses on the first attempt, and no longer leaves stale modal state after syncing.
- **Updater Stability:** Linux tray and background update checks now use GitHub release metadata instead of the native Electron updater, preventing crashes. Unsafe native updater paths are avoided on all platforms.
- **Linux Launcher Update:** `subminer -u` on Linux now performs release updates directly from the launcher without requiring the tray app to be running. When already on the latest version it reports up to date without downloading assets. Support asset updates are limited to the Linux rofi theme.
- **Linux Launcher Install:** First-run launcher installs on Linux now use a valid Bun shebang so the installed launcher executes correctly.
- **macOS Setup:** First-run setup now correctly recognizes launchers already installed via Homebrew or user PATH directories, and manual installs avoid writing to Homebrew-managed locations.
- **Update Dialog:** macOS update dialogs are brought to the front when `subminer --update` is run from the command line.
- **Setup Flow:** `subminer app --setup` now correctly opens the setup window when SubMiner is already running in the background. The standalone setup process also quits after first-run completes, returning the terminal prompt instead of leaving the app open.
- **Build:** One-shot `make clean build install` flows now correctly pick up the AppImage produced by the current build rather than a stale previous one.
- **Tray Settings:** Closing Yomitan settings launched from the tray no longer quits the tray app, and loading settings no longer blocks other tray actions. A close button is shown within the Yomitan settings page on Hyprland where native window controls are unavailable. The embedded Yomitan popup preview is disabled in the tray settings window to prevent renderer hangs. Extension refreshes are now serialized to prevent startup race conditions, and session help modals can close correctly without mpv running.

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
