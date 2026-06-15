> This is a prerelease build for testing. Stable changelog and docs-site updates remain pending until the final stable release.

## Highlights
### Added

- **Settings Window:** A dedicated Settings window is now available via `subminer --settings` or `subminer settings`, organized into Appearance, Behavior, Anki, Input, and Integration sections.
  - Includes click-to-learn keybinding controls, an AnkiConnect deck dropdown that auto-fills from Yomitan's current mining deck, and AnkiConnect-backed deck, field, and note-type pickers.
  - Live-saves changes for subtitle CSS declarations, stats keys, logging level, Anki field mappings, sentence card model, and other annotation and runtime options; search narrows across all categories including multi-word terms. AI and translation settings remain config-file only.

- **Auto-Updater:** SubMiner can now check for and apply updates from the system tray or by running `subminer -u`, with checksum verification and configurable update notifications.
  - The `subminer` launcher and Linux rofi theme update automatically alongside the app.
  - Set `updates.channel` to `"prerelease"` to receive beta and RC builds.

- **First-Run Setup:** A new optional setup flow installs Bun and the `subminer` command-line launcher on Linux, macOS, and Windows.
  - Windows users get a `subminer.cmd` PATH shim so `subminer` works in any terminal without manually adding `SubMiner.exe` to PATH.
  - Setup recognizes existing `subminer` installs in Homebrew or user PATH directories and avoids writing into Homebrew-owned paths. An Open SubMiner Settings button is included on completion; the standalone setup app quits after finishing.

- **Character Portraits:** Character-name subtitle matches can now show optional inline AniList character portraits.
  - Manual AniList title overrides are scoped per media directory so separate season folders keep independent character dictionary selections.

- **Log Export:** Sanitized log ZIP archives can be exported from the tray menu or by running `subminer logs -e`, with home-directory usernames redacted from the exported contents.

- **Logging Configuration:** SubMiner's logging level is now forwarded into launcher-started and Windows shortcut-started mpv sessions, controlling mpv log verbosity and plugin script logging.
  - The new `logging.rotation` config sets daily log retention (default 7 days). `logging.files` toggles let you enable or disable per-component log files; mpv logs are off by default unless explicitly enabled.

- **Yomitan Popup Visibility:** The new `subtitleStyle.primaryVisibleOnYomitanPopup` option keeps hover-mode primary subtitles visible while a Yomitan lookup popup is open.

- **Launcher:** `subminer --version` / `subminer -v` now prints the installed app version. The new `mpv.profile` config option passes an mpv profile to SubMiner-managed mpv launches, and bundled mpv plugin startup options are now configurable from SubMiner config.

### Changed

- **Subtitle Appearance:** Primary and secondary subtitle appearance now use color controls plus CSS declaration editors, saved as `subtitleStyle.css` and `subtitleStyle.secondary.css`; sidebar appearance uses `subtitleSidebar.css`.
  - Default font stack updated to `Hiragino Sans, M PLUS 1, Source Han Sans JP, Noto Sans CJK JP`; default text shadow is stronger, JLPT underlines are thicker, and the frequency `topX` threshold defaults to `10000`.
  - Existing configs are migrated automatically: legacy appearance options and hover token colors fold into `subtitleStyle.css`, and user config files are preserved.

- **Known-Word Colors:** Known-word and N+1 annotation colors moved to `subtitleStyle.knownWordColor` and `subtitleStyle.nPlusOneColor`. Legacy Anki color keys remain accepted with deprecation warnings.
  - N+1 highlighting is preserved for configs that already had it enabled; new configs leave it disabled unless `ankiConnect.nPlusOne.enabled` is set explicitly.

- **Character Dictionary:** Entries are now scoped to the current AniList media and generate Japanese name aliases only, so raw romanized or English aliases no longer appear as separate results.
  - A new `Ctrl/Cmd+D` manager modal lets you remove, reorder, or override loaded dictionary entries.
  - The in-app AniList title selector now waits for an explicit search rather than triggering automatically; the search box is prefilled from the current filename guess.

- **Linux Updater:** Tray "Check for Updates" now installs the new AppImage automatically via `electron-updater`, matching the macOS and Windows update flow. System-package-managed AppImages and non-AppImage launches fall back to the GitHub-asset flow.

- **Subsync:** The subtitle sync dialog now always opens the manual picker; the `subsync.defaultMode` config option has been removed.

- **Jellyfin Setup:** The server presets dropdown is replaced by a single editable server URL field.

- **Defaults:** Jellyfin remote-session startup warmup and character-name subtitle highlighting now default to off.

- **Runtime:** The bundled Electron runtime is updated from 39.8.6 to 42.2.0.

- **Subtitle Delay Shortcuts:** Default overlay subtitle delay and step bindings now match mpv's own defaults.
  - `z`, `Z`, and `x` adjust sub-delay; `Ctrl+Shift+Left/Right` runs native sub-step and shows the current delay on the OSD.
  - The previous SubMiner-only adjacent-cue delay action has been removed.

- **Update Notifications:** New installs default to overlay-only update notifications instead of overlay plus system notifications.

### Fixed

- **macOS Overlay:** Significantly improved overlay focus and stability across a range of scenarios.
  - The overlay hides when mpv loses focus, is minimized, or is no longer the foreground target; stays stable through transient window-tracking misses; remains correctly layered during stats mouse passthrough; and opens over fullscreen mpv without switching Spaces.
  - Passthrough is fixed so mpv controls stay clickable before hovering a subtitle bar. The compiled mpv window helper is now correctly bundled, preventing the overlay from falling back to a slower startup path on first launch.
  - Yomitan popup focus is restored after card mining or popup reload; clicking transparent overlay space correctly closes the popup and returns passthrough to mpv without a hide/reappear cycle.

- **Linux/Hyprland Overlay:** Overlay placement refreshes after leaving mpv fullscreen so the visible overlay stays aligned to the player.
  - The overlay stays stacked above mpv after click-to-focus events and is suspended while the in-player stats window is open.
  - Settings windows (SubMiner and Yomitan) now open above the subtitle overlay; the overlay hides immediately when the character dictionary modal opens, including while AniList lookup is in progress.
  - Auto-paused startup is more reliable: the overlay stays interactive during the first measurement gap, startup subtitle cache misses paint raw text before tokenization finishes, and temporarily empty mpv subtitle reads are resolved before synthetic warm readiness can resume playback.

- **Playback:** The first subtitle is primed before autoplay resumes so the overlay renders text before video playback begins. Launcher-owned videos quit SubMiner when playback ends while background and tray sessions stay alive.
  - The visible overlay and subtitle stream stay alive after restarting SubMiner from the `y-r` shortcut, with correct Linux bounds reapplication and user-paused playback preserved through readiness gates.
  - The visible overlay remains active when mpv advances to the next playlist item, even when the next episode loads after the warm transition delay.

- **Jellyfin Playback:** Resolved a wide range of discovery and playback issues: the active item is no longer reloaded during startup, paused mpv is no longer misreported as playing, startup unpause no longer repeats after a manual pause or `y-t` toggle, and duplicate ready signals no longer re-show the overlay.
  - Discovery now correctly handles delayed Japanese subtitle selection and prevents later-loading foreign tracks from stealing the active Japanese track.
  - Discovery resume correctly handles `StartPositionTicks: 0` for items with saved progress.

- **Jellyfin Subtitles:** Improved subtitle timing by preferring default embedded streams over external sidecars, stripping Jellyfin's server-selected stream from playback URLs, suppressing mpv auto-selection while SubMiner stages managed tracks, and automatically correcting Japanese-vs-English cue timeline offsets.
  - Per-stream subtitle delay shifts are restored on load. Track selection now tolerates transient `track-list` read failures and numeric string track IDs on Linux.

- **Jellyfin Overlay:** The visible subtitle overlay now shows automatically during Jellyfin playback so `subtitleStyle` appearance applies, and the bundled mpv plugin is injected when SubMiner auto-launches mpv so mpv-side keybindings work without overlay focus.
  - The `y-t` overlay toggle is reliable and remains sticky across stream redirects.
  - Passive Linux/Hyprland overlay shows no longer steal keyboard focus from mpv.

- **Jellyfin Remote Progress:** Fixed progress sync for mpv/SubMiner seek jumps, stopped sessions, startup path changes, and Linux websocket reconnect windows.
  - Play and Resume are now distinct: Play starts from the beginning while Resume starts at the saved position.
  - Final progress reports use SubMiner's last known position when mpv resets during stop.

- **Jellyfin Identity:** Cast device identity is now derived from the OS hostname. Multiple SubMiner installs no longer share the same remote-session identity.

- **Jellyfin Tray:** The discovery tray checkbox stays in sync on Linux after tray, CLI, or startup remote-session changes. Stale discovery sessions restart automatically when the server no longer lists the SubMiner cast target.

- **Jellyfin Setup:** Fixed the Windows login flow with an IPC bridge and immediate progress feedback; unreachable servers time out with an inline error instead of hanging.

- **AniList Progress:** Threshold checks now use fresh playback position data so updates fire correctly when playback reaches or skips past the watched threshold.
  - Season-specific results are preferred for multi-season files, with a clear message when the matched season is not in Planning or Watching status.
  - Repeated missing-token checks no longer exhaust AniList retry attempts or create duplicate dead-letter entries for the same episode.
  - Manual AniList linking from the stats anime page now strips generated `Season N` suffixes from automatic searches so results match the base title correctly.

- **Anki:** Sentence-audio padding is now opt-in by default; animated AVIF freeze-frame duration is correctly aligned to word audio length without double-counting padding.
  - Multi-line sentence mining stays aligned for repeated subtitle text; Kiku duplicate-card detection and merge flow are fixed; clipboard card updates from YouTube use mpv's resolved stream URLs; sentence cards refresh the secondary subtitle before saving.
  - Known-word cache append is fixed when no default Anki mining deck is configured but multiple known-word deck field mappings are present.

- **YouTube:** Primary subtitles are downloaded to temporary local files so the primary bar and sidebar read the same source, with cleanup on reload and quit.
  - False load-failure notifications are suppressed. Launcher-managed playback creates the tray icon when attaching to an already-running process, and app-owned playback no longer lets the mpv plugin start a second SubMiner instance.

- **Character Dictionary:** Surname honorifics are now matched for Japanese localized aliases embedded in AniList alternative names; cached snapshots are regenerated to include them.
  - Cached media matches are reused when loading a title with an existing snapshot, avoiding redundant AniList search requests. Manager keyboard shortcuts are correctly forwarded to the mpv plugin.

- **Updater:** Update checks are more stable across platforms: Linux uses GitHub release metadata; `subminer -u` can update independently of the tray app; macOS update dialogs reliably appear in the foreground.
  - Builds that cannot apply native updates show a manual-install message instead of a restart prompt. Windows retains the native NSIS update path while routing updater HTTP through the main process.
  - Linux updates now correctly create and refresh the launcher runtime plugin copy and rofi theme alongside AppImage and launcher updates; both support assets are auto-installed from the bundled app on first launch if either is missing.

- **Setup - macOS:** First-run setup recognizes existing `subminer` installs in Homebrew or user PATH directories and avoids writing into Homebrew-owned paths.
  - `subminer app --setup` opens the setup flow even when SubMiner is already running. The standalone setup app quits after completing first-run setup, and `subminer settings` exits cleanly when the window is closed.

- **Tray App:** Fixed several lifecycle issues: the tray stays running when Yomitan settings are closed; a close-only menu prevents accidentally quitting the tray app; an in-page close button is available on Hyprland where native window controls are unavailable.
  - Settings loading no longer blocks other tray actions; the embedded popup preview is disabled to prevent renderer hangs during sidebar navigation; extension refreshes at startup are serialized; the session help modal closes correctly without mpv running.
  - On Windows, "Open SubMiner Setup" now correctly opens the setup window after first-run setup is complete. System notifications now correctly show the SubMiner app icon when no custom notification image is provided.

- **Launcher:** Launcher-opened videos reuse an already-running background SubMiner instance and correctly reapply preferred subtitles on warm launches. Videos stay paused until subtitle priming and tokenization readiness complete.
  - `subminer settings` on macOS no longer emits Electron menu diagnostics and exits cleanly when the window is closed. Linux first-run launcher installs build with a valid Bun shebang; `subminer app` on Linux returns control to the terminal immediately.
  - On Windows, managed mpv launches from a background instance correctly retarget the new mpv socket, bind to the player window, and receive startup overlay options.

- **Subtitle Frequency:** Frequency highlighting is preserved for determiner-led noun compounds like `その場` while standalone determiners are still filtered. Annotations are corrected for Yomitan single-token compounds with internal particles like `目の前`.

- **Subtitle Annotation Prefetch:** Cached annotations and character images are ready for more live subtitle changes without delaying raw subtitle display.

- **Shortcuts:** Native mpv menu shortcuts are disabled during managed macOS playback so SubMiner shortcuts also work while mpv has focus. Session shortcuts including `stats.markWatchedKey` are correctly wired through mpv. The visible overlay receives focus when entering multi-line copy/mine selection so number keys work on macOS and Windows.

- **Stats:** In-player stats layering is fixed so delete confirmations, overlay modals, and update-check dialogs appear above the stats window. Jellyfin playback stats are grouped under item metadata so watched episodes merge with matching local library titles and display clean names.

- **Sidebar:** Yomitan lookup popups opened from the subtitle sidebar now correctly pause playback when popup auto-pause is enabled. Mined cards use audio and images from the clicked subtitle line rather than the current primary line.

- **Controller:** Config and debug shortcuts stay closed while controller support is disabled, with a notice to enable `controller.enabled`. Learn mode can be entered from the edit pencil or binding badge; remaps are saved per controller profile, and individual bindings can be reset to their defaults.

- **Discord Rich Presence:** Presence no longer falls back to Jellyfin stream URLs; Jellyfin playback titles are primed before loading tokenized streams so presence shows the show/episode title.

- **WebSocket:** The regular subtitle WebSocket now sends plain text only; annotation spans and token metadata are sent exclusively on the annotation WebSocket.

- **Windows Startup:** Fatal startup errors now show a native error dialog and write details to the app log instead of exiting silently.

- **Yomitan:** Fixed popups not opening when overlay startup races the Yomitan extension load.

- **Subtitle Sync Modal:** Fixed a macOS issue where the modal would flash and disappear on the first attempt, or leave stale state after syncing.

### Docs

- **Versioned Docs:** Stable docs are now published at the site root with current development docs under `/main/`.
  - Fixed versioned docs navigation so archived pages keep local links under the selected version, the version switcher no longer nests paths incorrectly, and local dev version routes serve warmed archive files instead of redirecting to production.

- **Configuration Reference:** All previously undocumented config options are now covered, including `subtitleStyle.primaryDefaultMode`, `stats.markWatchedKey`, `immersionTracking.lifetimeSummaries.*`, and all seven `mpv.*` launcher options. Updated known-word cache docs and examples to recommend expression/word fields.

- **Architecture Docs:** Added a Playback Startup Flow diagram and a Runtime Sockets section and diagram to the IPC + Runtime Contracts page, with cross-reference pointers in the MPV Plugin and Troubleshooting pages.

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
