> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

## Highlights
### Added

- **Settings Window:** A dedicated Settings window is now available via `subminer --settings` or `subminer settings`, organized into Appearance, Behavior, Anki, Input, and Integration sections. Includes click-to-learn keybinding controls, AnkiConnect-backed deck/field/note-type pickers, and live reload for stats keys, logging level, Jimaku, Subsync, YouTube language defaults, and Anki field mappings. AI and translation settings remain config-file only.

- **Auto-Updater:** SubMiner can now check for and apply updates from the system tray or by running `subminer -u`, with checksum verification, configurable update notifications, and an opt-in prerelease channel. The `subminer` launcher and Linux rofi theme update automatically. Set `updates.channel` to `"prerelease"` to receive beta and RC builds.

- **First-Run Setup:** A new optional setup flow installs Bun and the `subminer` command-line launcher on Linux, macOS, and Windows, with an Open SubMiner Settings button on completion. Windows users get a `subminer.cmd` PATH shim so `subminer` works in any terminal without manually adding `SubMiner.exe` to PATH.

- **Launcher:** `subminer --version` / `subminer -v` now prints the installed app version. The new `mpv.profile` config option passes an mpv profile to SubMiner-managed mpv launches. Bundled mpv plugin startup options are now configurable from SubMiner config.

- **Character Portraits:** Character-name subtitle matches can now show optional inline AniList character portraits. Manual AniList title overrides are scoped per media directory so separate season folders keep independent character dictionary selections.

### Changed

- **Subtitle Appearance:** Primary and secondary subtitle appearance now use color controls plus CSS declaration editors, saved as `subtitleStyle.css` and `subtitleStyle.secondary.css`. Sidebar appearance is configured via `subtitleSidebar.css`. The default subtitle font stack is updated to `Hiragino Sans, M PLUS 1, Source Han Sans JP, Noto Sans CJK JP`. Existing configs are migrated automatically.

- **Known-Word Colors:** Known-word and N+1 annotation colors moved to `subtitleStyle.knownWordColor` and `subtitleStyle.nPlusOneColor`. Legacy Anki color keys remain accepted with deprecation warnings. N+1 highlighting is preserved for configs that already had it enabled; new configs leave it disabled unless `ankiConnect.nPlusOne.enabled` is set explicitly.

- **Character Dictionary:** A new `Ctrl/Cmd+D` manager modal lets you remove, reorder, or override loaded dictionary entries. The in-app AniList title selector now waits for an explicit search rather than triggering automatically; the search box is prefilled from the current filename guess so it can be edited before confirming an override. Lookup entries are scoped to generated Japanese name aliases only, so raw romanized or English aliases no longer appear as separate results.

- **Linux Updater:** Tray "Check for Updates" now installs the new AppImage automatically via `electron-updater`, matching the macOS and Windows update flow. System-package-managed AppImages (e.g. AUR `/opt/SubMiner`) and non-AppImage launches fall back to the GitHub-asset flow.

- **Subsync:** The subtitle sync dialog now always opens the manual picker; the `subsync.defaultMode` config option has been removed.

- **Jellyfin:** The server presets dropdown in Jellyfin setup is replaced by a single editable server URL field.

- **AniSkip:** The key binding setting now uses click-to-learn key capture instead of raw text entry.

- **Setup:** The bundled mpv runtime plugin readiness card is removed from first-run setup; the legacy mpv plugin removal notice still appears when needed.

- **Defaults:** Jellyfin remote-session startup warmup and character-name subtitle highlighting now default to off.

- **Runtime:** The bundled Electron runtime is updated from 39.8.6 to 42.2.0.

### Fixed

- **macOS Overlay:** Significantly improved overlay focus and stability: the overlay hides when mpv loses focus, is minimized, or is no longer the foreground target; stays stable through transient window-tracking misses; remains correctly layered during stats mouse passthrough; opens over fullscreen mpv without switching Spaces; and stays stable when mpv remains frontmost but window geometry temporarily disappears from macOS APIs. Passthrough is fixed so mpv controls stay clickable before hovering a subtitle bar. The overlay stays stable when clicking from the overlay back into mpv. Background tracking overhead is reduced while mpv is stably focused.

- **Linux/Hyprland Overlay:** Overlay placement refreshes after leaving mpv fullscreen so the visible overlay stays aligned to the player. The visible overlay remains stacked above mpv after mpv regains focus from clicks, and is suspended while the in-player stats window is open. Settings windows (SubMiner and Yomitan) now open above the subtitle overlay on Hyprland instead of behind it.

- **Jellyfin Playback:** Resolved a wide range of Jellyfin discovery and playback issues: the active item is no longer reloaded during startup, paused mpv is no longer misreported as playing, startup unpause no longer repeats after a manual pause or `y-t` toggle, duplicate ready signals no longer re-show the overlay, and long-lived sidebar ffmpeg extractors no longer run against stream URLs. Discovery now correctly handles delayed Japanese subtitle selection and prevents later-loading foreign tracks from stealing the active Japanese track. Discovery resume correctly handles `StartPositionTicks: 0` for items with saved progress.

- **Jellyfin Subtitles:** Improved subtitle timing by preferring default embedded streams over external sidecars, stripping Jellyfin's server-selected stream from playback URLs, suppressing mpv auto-selection while SubMiner stages managed tracks, and automatically correcting clear Japanese-vs-English cue timeline offsets. Per-stream subtitle delay shifts are restored on load. Track selection now tolerates transient `track-list` read failures and numeric string track IDs on Linux.

- **Jellyfin Overlay:** The visible subtitle overlay now shows automatically during Jellyfin playback so `subtitleStyle` appearance applies. The bundled mpv plugin is injected when SubMiner auto-launches mpv for Jellyfin so mpv-side keybindings work without overlay focus. The `y-t` overlay toggle is reliable and remains sticky across stream redirects. Passive Linux/Hyprland overlay shows no longer steal keyboard focus from mpv.

- **Jellyfin Remote Progress:** Fixed progress sync for mpv/SubMiner seek jumps, stopped sessions, startup path changes, and Linux websocket reconnect windows. Play and Resume are now distinct: Play starts from the beginning while Resume starts at the saved position. Final progress reports use SubMiner's last known position when mpv resets during stop.

- **Jellyfin Identity:** Cast device identity is now derived from the OS hostname. Multiple SubMiner installs no longer share the same remote-session identity, and SubMiner always reports itself as the client regardless of legacy configurable identity fields.

- **Jellyfin Tray:** The discovery tray checkbox stays in sync on Linux after tray, CLI, or startup remote-session changes. Stale discovery sessions restart automatically when the server no longer lists the SubMiner cast target. Library discovery works correctly when the app log level is set above info.

- **Jellyfin Setup:** Fixed the Jellyfin setup login flow on Windows: login now uses an IPC bridge with immediate progress feedback, and unreachable servers time out with an inline error instead of hanging.

- **Subtitle Sync Modal:** Fixed a macOS issue where opening the subtitle sync modal would flash and disappear on the first attempt, or leave stale state after syncing.

- **Controller:** Controller config and debug shortcuts now stay closed while controller support is disabled, with a notice to enable `controller.enabled`. Learn mode can be entered from the edit pencil or binding badge, remaps are saved per controller profile, and individual bindings can be reset to their defaults.

- **AniList Progress:** Progress threshold checks now use fresh playback position data so updates fire correctly when playback reaches or skips past the watched threshold. Season-specific results are preferred for multi-season files, and a clear message is shown when the matched season is not in Planning or Watching status.

- **Anki:** Sentence-audio padding is now opt-in by default. When padding is configured, animated AVIF freeze-frame duration is correctly aligned to the word audio length without double-counting sentence padding. Multi-line sentence mining stays aligned when repeated subtitle text appears in the selected history range. Manual clipboard card updates from YouTube playback now use mpv's resolved stream URLs for generated audio and images.

- **YouTube:** Primary subtitles are now downloaded to temporary local files so the primary bar and sidebar read the same source, with cleanup on reload and quit. False subtitle load failure notifications are suppressed after SubMiner confirms the selected track loaded. Launcher-managed playback commands create the tray icon even when attaching to an already-running process, and app-owned YouTube playback no longer lets the mpv plugin start a second SubMiner instance.

- **Character Dictionary:** Cached media matches are reused when loading a title with an existing snapshot, avoiding redundant AniList search requests on repeat visits. The visible subtitle overlay is suppressed as soon as the character dictionary modal opens, including while AniList lookup is loading or returns no results.

- **Updater:** Update checks are more stable across platforms: Linux uses GitHub release metadata instead of the native Electron updater; `subminer -u` can update independently of the tray app; macOS update dialogs reliably appear in the foreground; builds that cannot apply native updates show a manual-install message instead of a restart prompt; Windows retains the native NSIS update path while routing updater HTTP through the main process; and macOS updater metadata mismatches from conflicting ZIP filenames are resolved.

- **Setup - macOS:** First-run setup now recognizes existing `subminer` installs in Homebrew or user PATH directories, and manual setup avoids writing into Homebrew-owned paths. `subminer app --setup` opens the setup flow even when SubMiner is already running in the background. The standalone setup app quits after completing first-run setup, and `subminer settings` exits cleanly when the window is closed.

- **Tray App:** Fixed several lifecycle issues with tray-launched Yomitan settings: the tray stays running when settings are closed; settings loading no longer blocks other tray actions; a close-only menu prevents accidentally quitting the tray app; an in-page close button is available on Hyprland where native window controls are unavailable; the embedded popup preview is disabled to prevent renderer hangs during sidebar navigation; extension refreshes at startup are serialized to prevent race conditions; and the session help modal can close correctly without mpv running. On Windows, the tray "Open SubMiner Setup" action now correctly opens the setup window after first-run setup is complete.

- **Launcher:** Launcher-opened videos reuse an already-running background SubMiner instance and correctly reapply preferred subtitles on warm launches. Videos stay paused when attaching to a running background app until subtitle priming and tokenization readiness complete. Launcher-owned tray apps close after playback ends. `subminer settings` on macOS no longer emits Electron menu diagnostics. Linux first-run launcher installs now build with a valid Bun shebang. `subminer app` on Linux returns control to the terminal immediately. On Windows, managed mpv launches from a background SubMiner instance correctly retarget the new mpv socket, bind to the player window, and receive startup overlay options.

- **Playback:** The first subtitle is primed before autoplay resumes so the overlay renders text before video playback begins. Launcher-owned videos quit SubMiner when playback ends while background and tray sessions stay alive.

- **Subtitle Frequency:** Frequency highlighting is preserved for determiner-led noun compounds like `その場` while standalone determiners are still filtered.

- **Shortcuts:** Native mpv menu shortcuts are disabled during managed macOS playback so configured SubMiner shortcuts also work while mpv has focus. Session shortcuts including `stats.markWatchedKey` are correctly wired through mpv. The visible overlay receives focus when entering multi-line copy/mine selection so number keys work on macOS and Windows.

- **Overlay Restart:** The visible overlay and subtitle stream stay alive after restarting SubMiner from the `y-r` shortcut, with correct bounds reapplication on Linux and user-paused playback preserved through readiness gates.

- **Stats:** In-player stats layering is fixed so delete confirmations, overlay modals, and update-check dialogs appear above the stats window. Jellyfin playback stats are grouped under item metadata instead of stream URLs, so watched episodes merge with matching local library titles and display clean names.

- **Sidebar:** Yomitan lookup popups opened from the subtitle sidebar now correctly pause playback when popup auto-pause is enabled.

- **Discord Rich Presence:** Presence no longer falls back to Jellyfin stream URLs; Jellyfin playback titles are primed before loading tokenized streams so presence shows the show/episode title.

- **WebSocket:** The regular subtitle WebSocket now sends plain text only; annotation spans and token metadata are sent exclusively on the annotation WebSocket.

- **Windows:** Startup failures now show a native error dialog and write fatal details to the app log instead of exiting silently.

- **Yomitan:** Fixed Yomitan popups not opening when overlay startup races the Yomitan extension load.

- **Settings:** Search now works across all categories, narrows correctly on multi-word terms, and hides settings with dedicated editors. Live saves for subtitle CSS declarations apply immediately to open overlays. Legacy subtitle appearance options and hover token colors are automatically migrated into `subtitleStyle.css`. The note-fields note type picker defaults to the configured Anki deck's note type, then `Kiku`, then `Lapis`, leaving it blank for manual selection otherwise. User config files are preserved during legacy config compatibility handling. The generated example config uses the same CSS declaration paths written by the Settings window.

### Docs

- **Versioned Docs:** Stable docs are now published at the site root with current development docs under `/main/`. Fixed versioned docs navigation so archived pages keep local links under the selected version, the version switcher no longer nests paths incorrectly, local dev version routes serve warmed archive files instead of redirecting to production, and internal README files no longer break archived builds.

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
