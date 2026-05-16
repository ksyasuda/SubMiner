> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

## Highlights
### Added

**Auto-Update:** The tray and `subminer -u` now check for SubMiner releases and prompt you to install updates, covering the app itself, the command-line launcher, and Linux rofi themes. Update notifications are configurable, downloads are checksum-verified, and an opt-in prerelease channel lets you receive beta and RC builds.

**First-Run Setup:** The app can now install Bun and the `subminer` command-line launcher during first run on Linux, macOS, and Windows. On Windows, a `subminer.cmd` PATH shim lets you type `subminer` in any terminal without manually adding `SubMiner.exe` to PATH.

### Fixed

**macOS Overlay:** Controls on the mpv window are now clickable before you hover the subtitle bars. Transient mpv window focus misses no longer hide the overlay; minimizing mpv still hides it as expected.

**Subtitle Sync:** The subtitle sync modal on macOS now opens reliably: no more flash-and-hide on the first attempt or stale modal state left after syncing.

**Updater:** Update checks on Linux now use GitHub release metadata instead of the native Electron updater, preventing tray app crashes. macOS builds that cannot auto-install updates show a manual install prompt instead of a non-functional restart prompt. macOS update dialogs are brought to the front when `subminer --update` is run from the launcher.

**Linux Command-Line Updater:** `subminer -u` now performs release updates directly from the launcher without requiring the tray app, and reports "up to date" without downloading assets when already on the latest version.

**Launcher Setup:** Linux first-run launcher installs now build with a valid Bun shebang. `subminer app --setup` correctly opens the setup flow even when SubMiner is already running in the background. On macOS, first-run setup recognizes existing launchers in Homebrew or user PATH directories, and manual installs avoid Homebrew-managed locations. The setup window now quits automatically after first-run setup completes, returning control to the terminal.

**Tray App:** Fixed several issues with tray-launched Yomitan settings: closing the window no longer quits the tray app, a close-only menu replaces the default native menu, and settings loading no longer blocks other tray actions. An in-page close button is now available on Hyprland, where native window controls may be absent. Disabled the embedded popup preview in the settings window to prevent renderer hangs. Fixed a startup race condition that could leave extension loading in an error state, and corrected focus handling for the session help modal when mpv is not running.

**Build:** One-shot `make clean build install` flows now correctly pick up the AppImage built in the same invocation.

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
