## Highlights
### Breaking Changes

- **Subsync:**
  - The `subsync.defaultMode` config option has been removed
  - Subsync now always opens the manual subtitle picker regardless of any previously set default mode
- **N+1 Highlighting:**
  - N+1 highlighting now has its own dedicated `ankiConnect.nPlusOne.enabled` option, separate from known-word highlighting
  - It is no longer enabled automatically when known-word highlighting is on — enable it explicitly to keep N+1 annotations

### Added

- **Auto-Updater:**
  - Tray and `subminer -u` update checks with app update prompts
  - Launcher and Linux rofi theme auto-updates
  - Checksum verification and configurable notifications
  - Opt-in prerelease channel via `updates.channel: "prerelease"`
- **Settings Window:**
  - New dedicated Settings window via `subminer --settings` or `subminer settings`, organized into Appearance, Behavior, Anki, Input, and Integration sections
  - Click-to-learn keybinding controls
  - AnkiConnect-backed deck, field, and note-type pickers that auto-fill from the configured Anki deck
  - Cross-category search
  - Live save for most options including subtitle CSS, stats keys, logging level, Jimaku, Subsync, and Anki mappings
  - AI and translation settings remain config-file only
- **Inline Character Portraits:**
  - Optional AniList character portraits appear inline for name-matched subtitle text
  - Manual AniList overrides scoped per parent media directory so separate season folders maintain separate character dictionary selections
- **Character Dictionary Manager:** New `Ctrl/Cmd+D` manager modal to remove, reorder, or override loaded entries.
- **Log Export:** Sanitized log ZIP export from the tray menu and via `subminer logs -e`, with home-directory usernames redacted from exported contents.
- **Launcher CLI:**
  - `subminer --version` / `subminer -v` prints the installed app version
  - `mpv.profile` config and Settings support passes a named mpv profile to managed launches
  - Bundled mpv plugin startup options are now configurable from SubMiner config
- **First-Run Setup:**
  - Optional installer for Bun and the `subminer` CLI on Linux, macOS, and Windows
  - Windows `subminer.cmd` PATH shim so `subminer` works without manually adding `SubMiner.exe` to PATH
  - Setup recognizes existing Homebrew or user PATH installs and avoids writing into Homebrew-owned paths
  - Includes an Open SubMiner Settings button
  - Standalone setup app quits after completing, returning terminal control
- **Primary Subtitle Visibility on Yomitan Popup:** New `subtitleStyle.primaryVisibleOnYomitanPopup` option keeps hover-mode primary subtitles visible while a Yomitan popup is open.

### Changed

- **Subtitle Appearance Config:**
  - Primary and secondary subtitle appearance now use color controls plus CSS declaration editors, saved as `subtitleStyle.css`, `subtitleStyle.secondary.css`, and `subtitleSidebar.css`
  - Known-word and N+1 annotation colors moved to `subtitleStyle.knownWordColor` and `subtitleStyle.nPlusOneColor`
  - Subtitle font defaults updated to `Hiragino Sans, M PLUS 1, Source Han Sans JP, Noto Sans CJK JP`
  - Existing configs migrate automatically; legacy Anki color keys still accepted with deprecation warnings
- **Subtitle Style Defaults:**
  - Stronger outline-style text shadow
  - Thicker JLPT underlines
  - Frequency `topX` default raised to `10000`
- **Character Dictionary:**
  - Entries scoped to the current AniList media for name matching and inline portraits
  - Generates Japanese-only name aliases so raw romanized/English aliases no longer surface as separate results
  - In-app AniList selector waits for an explicit search with the box prefilled from the current filename
  - `subtitleStyle.nameMatchEnabled` is now the sole switch for dictionary sync and builds
- **Electron Runtime:** Updated from 39.8.6 to 42.2.0, returning SubMiner to a supported Electron release line.
- **Jellyfin Setup:**
  - Removed the server presets dropdown
  - Setup now shows a single editable server URL field
- **Jellyfin Cast Identity:**
  - Device identity now derived from the OS hostname and always reported as SubMiner
  - Previously configurable identity fields are ignored, preventing multiple installs from sharing a remote-session identity
- **Startup Defaults:** Jellyfin remote-session startup warmup and character-name subtitle highlighting now default to off.
- **Setup Appearance:** Removed the bundled mpv runtime plugin readiness card from the setup flow.

### Fixed

- **AniList Progress:**
  - Progress updates fire correctly when playback reaches or skips past the watched threshold, using fresh mpv timing events
  - Season-specific results preferred for multi-season files, with a clear message when the matched season is not in Planning or Watching
  - Repeated missing-token checks no longer exhaust retry attempts or duplicate dead-letter entries
- **Anki Mining:**
  - Sentence-audio padding is opt-in by default
  - Animated AVIF freeze-frame duration aligned to word audio length without double-counting
  - Multi-line sentence alignment fixed for repeated subtitle text
  - Kiku duplicate-card detection, auto-merge, modal acknowledgment race, and field/tag ordering corrected
  - YouTube playback cards use mpv's resolved stream URLs
  - Sentence cards refresh the secondary subtitle before saving
- **Jellyfin Discovery:**
  - Startup, subtitle track selection, and duplicate ready-signal handling all fixed
  - Paused mpv no longer misreported as playing
  - Resume corrected when a remote play command sends `StartPositionTicks: 0` despite saved progress
- **Jellyfin Remote:**
  - Tray checkbox stays in sync on Linux after tray, CLI, or startup changes
  - Remote controller visibility and progress sync fixed for seeks, stops, startup path changes, and Linux websocket reconnect windows
  - Play and Resume now behave correctly (Play from beginning, Resume from saved position)
  - Final progress reports reuse SubMiner's last known position when mpv resets on stop
  - Windows setup login flow fixed with an IPC bridge, immediate feedback, and a timeout with inline error for unreachable servers
- **Overlay (macOS):**
  - Overlay hides when mpv loses focus, is minimized, or is no longer the foreground app
  - Stays stable through transient window geometry disappearances from macOS APIs and when clicking from the overlay back into mpv
  - Stats overlay opened inactive so it appears over fullscreen mpv without switching Spaces
  - Passthrough fixed so mpv controls stay clickable before hovering a subtitle bar
- **Yomitan Sidebar:**
  - Playback stays paused for sidebar-opened Yomitan popups when auto-pause is enabled
  - Popups now open when startup races the Yomitan extension load
  - Sidebar mining cards use audio and images from the clicked sidebar line instead of the current primary subtitle
- **Launcher:**
  - `subminer app` on Linux returns terminal control immediately
  - `subminer app --setup` opens the setup flow when SubMiner is already running in the background
- **YouTube Playback:**
  - Selected subtitles downloaded to local temp files so the primary bar and sidebar read the same source, with cleanup on reload and quit
  - False load-failure notifications suppressed
  - Tray icon created on launcher-managed playback that attaches to an already-running process
- **Shortcuts:**
  - Native mpv menu shortcuts disabled during managed macOS playback so configured SubMiner shortcuts work while mpv has focus
  - Custom session shortcuts including `stats.markWatchedKey` wired through mpv
  - Multi-line copy/mine overlay correctly focused so number keys choose the line count on macOS and Windows
- **Controller Bindings:**
  - Controller config and debug shortcuts stay closed while controller support is disabled
  - Binding learn mode starts from the edit pencil
  - Remaps saved per controller profile
  - Binding badges also start learn mode
  - Row reset buttons restore individual bindings to defaults
- **Logging:**
  - `logging.level` forwarded to launcher-started and Windows shortcut-started mpv sessions, covering mpv log verbosity, plugin logging, and plugin-launched app logging
  - `logging.rotation` (default 7 days) and per-component `logging.files` toggles added, with mpv logs disabled by default
  - Repeated IPC socket warning spam suppressed while waiting for mpv to recreate the socket
  - Windows mpv IPC, subtitle track, and Yomitan diagnostics added
- **In-Player Stats:**
  - Layering fixed so delete confirmations, overlay modals, and update-check dialogs appear above the stats window
  - Jellyfin playback stats grouped by item metadata so watched episodes merge with matching local library titles and keep clean display names
- **WebSocket Annotations:**
  - Annotation spans and token metadata stay on the annotation WebSocket
  - The regular subtitle WebSocket is plain-text only
- **Subtitle Annotation Prefetching:** Cached colored annotations and character images ready sooner for live subtitle changes without delaying raw subtitle display.
- **Windows Startup Errors:** Fatal startup failures now show a native error dialog and write details to the app log instead of exiting silently.

### Docs

- **Documentation Site:**
  - Published stable docs at the site root with current development docs under `/main/`
  - Fixed versioned docs navigation, archived page link handling, and local dev version routing
  - Documented all previously undocumented config options including `subtitleStyle.primaryDefaultMode`, `stats.markWatchedKey`, `immersionTracking.lifetimeSummaries.*`, and all seven `mpv.*` launcher options
  - Added Playback Startup Flow and Runtime Sockets diagrams to the architecture docs with cross-reference pointers in the MPV Plugin and Troubleshooting pages

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Windows: `SubMiner-*.exe` and `SubMiner-*-win.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
