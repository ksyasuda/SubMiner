> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

## Highlights
### Added

- **Settings Window:** A dedicated Settings window is now available via `subminer --settings` or `subminer settings`. Options are organized into Appearance, Behavior, Anki, Input, and Integration sections with learned keybinding controls, AnkiConnect-backed deck/field/note-type pickers, and live reload for stats keys, logging level, Jimaku, Subsync, YouTube language defaults, and Anki field mappings. AI and translation fields remain supported in config files only.

- **Auto-Updater:** SubMiner can now check for and apply updates from the system tray or by running `subminer -u`. Checks include checksum verification, configurable notifications, and an opt-in channel for prerelease builds. The `subminer` launcher and Linux rofi theme are also updated automatically.

- **First-Run Setup:** A new optional setup flow installs Bun and the `subminer` command-line launcher on Linux, macOS, and Windows. Windows users get a `subminer.cmd` PATH shim so `subminer` works in any terminal without manually adding `SubMiner.exe` to PATH. First-run setup also includes an Open SubMiner Settings button.

- **Launcher:** `subminer --version` / `subminer -v` now prints the installed SubMiner app version.

### Changed

- **Subtitle Appearance:** Primary and secondary subtitle appearance now use color controls plus CSS declaration editors, saved as `subtitleStyle.css` and `subtitleStyle.secondary.css`. Existing configs are migrated automatically. Sidebar appearance is now configured via `subtitleSidebar.css`; the default subtitle font is updated to `Hiragino Sans, M PLUS 1, Source Han Sans JP, Noto Sans CJK JP`.

- **Known-Word Colors:** Known-word and N+1 annotation colors moved to `subtitleStyle.knownWordColor` and `subtitleStyle.nPlusOneColor`. Legacy Anki color keys are still accepted with deprecation warnings. Existing configs that had known-word highlighting enabled retain N+1 highlighting; new configs leave N+1 disabled unless `ankiConnect.nPlusOne.enabled` is explicitly set.

- **Linux Updater:** Tray "Check for Updates" now automatically installs the new AppImage via `electron-updater`, matching the macOS and Windows tray flow. AppImages managed by a system package (e.g. AUR `/opt/SubMiner`) and non-AppImage launches fall back to the GitHub-asset flow.

- **Jellyfin:** The server presets dropdown in Jellyfin setup is removed; setup now shows a single editable server URL field.

- **AniSkip:** The key binding setting now uses click-to-learn key capture instead of raw text entry.

- **Setup:** The bundled mpv runtime plugin readiness card is removed from first-run setup; the legacy mpv plugin removal notice still appears when needed.

- **Defaults:** Jellyfin remote-session startup warmup and character-name subtitle highlighting now default to off.

### Fixed

- **macOS Overlay:** Significantly improved overlay focus and stability: the overlay now hides when mpv loses focus or is minimized, stays stable through transient window-tracking misses, remains correctly layered during stats mouse passthrough, and opens over fullscreen mpv without switching Spaces. Passthrough is fixed so mpv controls stay clickable before hovering a subtitle bar. Background tracking overhead is reduced while mpv is stably focused.

- **Subtitle Sync Modal:** Fixed a macOS issue where opening the subtitle sync modal would flash and disappear on the first attempt, or leave stale state after syncing.

- **Controller:** Controller config and debug shortcuts now stay closed while controller support is disabled, with a notice to enable `controller.enabled`. Learn mode can be entered from the edit pencil or binding badge, remaps are saved per controller profile, and individual bindings can be reset to their defaults.

- **AniList Progress:** Progress threshold checks now use fresh playback position data so updates fire correctly when playback reaches or skips past the watched threshold. Season-specific results are preferred for multi-season files, and a clear message is shown when the matched season is not in Planning or Watching status.

- **Character Dictionary:** Cached media matches are reused when loading a title with an existing snapshot, avoiding redundant AniList search requests on repeat visits.

- **Updater:** Update checks are more stable across platforms: Linux uses GitHub release metadata instead of the native Electron updater, `subminer -u` can update independently of the tray app, macOS update dialogs come to the front reliably, unsupported builds show a manual-install message, and Windows keeps the native NSIS update path while routing updater HTTP through the main process. GitHub release lookups now avoid Electron networking on Linux and macOS.

- **Setup - macOS:** First-run setup now recognizes existing `subminer` launcher installs in Homebrew or user PATH directories, and manual setup avoids writing into Homebrew-owned paths. `subminer app --setup` opens the setup flow even when SubMiner is already running in the background. The standalone setup app quits after completing first-run setup, and `subminer settings` exits cleanly when the window is closed - both return control to the terminal without requiring Ctrl+C.

- **Tray App:** Fixed several lifecycle issues with tray-launched Yomitan settings: the tray stays running when settings are closed; settings loading no longer blocks other tray actions; the settings window uses a close-only menu to prevent accidentally quitting the tray app; an in-page close button is provided on Hyprland where native window controls are unavailable; the embedded popup preview is disabled to prevent renderer hangs during sidebar navigation; extension refreshes at startup are serialized to prevent race conditions; and the session help modal can now close correctly without mpv running.

- **Launcher - Linux:** First-run launcher installs are now built with a valid Bun shebang, fixing installs that previously failed silently.

- **Launcher:** Launcher-opened videos now reuse an already-running background SubMiner instance and correctly reapply preferred subtitles on warm launches. Videos stay paused when attaching to a running background app until subtitle priming and tokenization readiness complete. Launcher-owned tray apps close after playback ends.

- **Playback:** The first subtitle is primed before autoplay resumes so the overlay can render text before video playback begins.

- **Subtitle Frequency:** Frequency highlighting is preserved for determiner-led noun compounds like `その場` while standalone determiners are still filtered.

- **macOS Shortcuts:** Native mpv menu shortcuts are disabled during managed macOS playback so configured SubMiner shortcuts work while mpv has focus. Session shortcuts including `stats.markWatchedKey` are now correctly wired through mpv.

- **Overlay Restart:** The visible overlay and subtitle stream now stay alive after restarting SubMiner from the `y-r` shortcut, with correct bounds reapplication on Linux and user-paused playback preserved through readiness gates.

- **WebSocket:** The subtitle WebSocket is now plain-text only; annotation spans and token metadata are sent exclusively on the annotation WebSocket.

- **Jellyfin:** Fixed the setup popup login path on Windows using an IPC bridge, with immediate login progress feedback and a timeout for unreachable server attempts.

- **Windows:** Startup failures now show a native error dialog and write fatal details to the app log instead of exiting silently.

- **Yomitan:** Fixed Yomitan popups not opening when overlay startup races the Yomitan extension load.

- **Settings:** Settings window search now searches across all categories, narrows correctly on multi-word terms, and hides settings with dedicated editors. Live saves for subtitle CSS declarations apply immediately to open overlays. Legacy subtitle appearance options and hover token colors are automatically migrated into `subtitleStyle.css`.

- **Build - Linux:** Fixed one-shot `make clean build install` flows so the install step correctly picks up the AppImage produced earlier in the same make invocation.

### Docs

- **Versioned Docs:** Stable docs are now published at the site root with current development docs under `/main/`. Fixed versioned docs navigation so archived pages keep local links under the selected version, the version switcher no longer nests paths incorrectly, local dev version routes serve warmed archive files instead of redirecting to production, and internal README files no longer break archived builds.

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
