# Changelog

## v0.15.0 (2026-05-29)

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

<details>
<summary>Internal changes</summary>

### Internal

- **Release Tooling:**
  - Release-note polishing treats pending fragments and reviewed prerelease notes as a cumulative final outcome, collapsing prerelease-only fixes into the final user-facing change
  - Prerelease generation reuses existing reviewed notes and merges only new fragment material
  - `make clean` preserves `release/prerelease-notes.md`
- **Tests:** Removed stale Yomitan vendor source-inspection assertions for changes that were not shipped.

</details>

## v0.14.0 (2026-05-12)

### Added

- **Character Dictionary:** Added AniList-based character dictionary selection for resolving title mismatches — open it in-app with the new `Ctrl+Alt+A` shortcut or from the CLI with `subminer dictionary --candidates` / `--select`. Series-scoped overrides replace stale entries in the merged dictionary.
- **Primary Subtitle Bar:** Added a `V` shortcut and mpv plugin binding to toggle the primary subtitle bar without affecting mpv's native subtitle visibility.
- **Texthooker:** Added `subminer texthooker -o` and a tray menu item to open the local texthooker page in the default browser.

### Changed

- **mpv Plugin Setup:** SubMiner-managed launches now automatically inject the bundled mpv plugin when no global plugin is installed. The legacy plugin removal prompt appears before mpv starts so users can trash the old global plugin and relaunch immediately.
- **Tray Menu:** Replaced the "Open Overlay" menu item with "Open Help", which opens the session help modal.
- **Stats Exclusions:** Vocabulary exclusions now persist in the immersion database and import any existing browser-local exclusions on first load.
- **Config Example:** The generated example config now lists every built-in default keybinding, with regression coverage ensuring they compile, dispatch in the overlay, and register through the mpv plugin.
- **Default Startup Options:** Texthooker startup and subtitle/annotation WebSocket servers are now disabled by default. Fresh installs also get updated primary subtitle style defaults: a Japanese font stack, transparent subtitle backgrounds, stronger text shadows, and teal N4/fourth-band coloring.

### Fixed

- **JLPT Underlines:** Restored persistent JLPT underlines on subtitle tokens; they now keep their JLPT color after dictionary lookups and when a token also carries a known-word or frequency annotation color.
- **Hyprland Fullscreen Overlay:** Fixed fullscreen transitions on Hyprland so mpv fullscreen changes refresh overlay geometry, reassert topmost stacking, and keep hover-pause working through resize/toggle cycles. The overlay now aligns exactly with the mpv window, disables floating-window decoration on exact placement, and no longer pins across workspaces.
- **macOS Overlay:** Kept the overlay visible and interactive during transient tracker refreshes and while mpv is the active tracked window. Fixed z-order so the overlay stays above mpv but behind other foreground windows.
- **Subtitle Annotation Colors:** Fixed annotated token colors so `subtitleStyle` typography is preserved and higher-priority known-word and frequency colors correctly take precedence over JLPT colors. Added a subtle brightness lift for hover states so transparent hover backgrounds still show a visible affordance. Hid the browser focus ring on the overlay surface so focused overlays no longer show a viewport border.
- **Grammar and Particle Annotation Filters:** Suppressed annotation styling (N+1, JLPT, frequency, name) for grammar-only tokens including auxiliary inflection fragments like `れる`/`れた`, polite copula tails like `です`/`ですよ`, negative copula phrases like `じゃないですか`, and standalone particles like `に`; lexical forms like `くれる` remain eligible.
- **Interjection and Existence Verb Filters:** Suppressed annotations for ハァ-style interjections and `ある`/`有る` existence verbs; known-word highlighting still applies for `ある`.
- **Known-Word Compound Tokens:** Preserved Yomitan compound token boundaries for known-word highlighting so known component words no longer color a larger unknown compound green.
- **Kana Token Annotations:** Stopped kana-only tokens from being selected as N+1 targets or receiving N+1, JLPT, or frequency annotation metadata.
- **Subtitle Bar Modes:** Changed `v` to cycle the primary subtitle bar through visible, hover, and hidden modes with OSD feedback. A new `subtitleStyle.primaryDefaultMode` config option sets the startup default independently of secondary subtitles.
- **Subtitle Keybindings:** Fixed the default replay/next subtitle keybindings by moving session help to `Ctrl/Cmd+/`, freeing `Ctrl+Shift+H` and `Ctrl+Shift+L` for subtitle playback controls. The mpv plugin now registers shifted letter chords correctly so `Ctrl+Shift+L` reaches play-next-subtitle; play-next also starts playback from a paused state before re-pausing at subtitle end.
- **Yomitan Popup Precedence:** Fixed keyboard-only Yomitan popup controls so popup shortcuts take precedence over overlay keybindings like `j`.
- **Subtitle Prefetch:** Kept annotation prefetch running after immediate cache-hit renders so upcoming subtitle colors stay ready.
- **Frequency Highlighting:** Fixed frequency highlighting for ordinal prefix-noun tokens like `第二` so JPDB ranks are preserved in subtitle annotations.
- **Known-Word Refresh After Mining:** Refreshed the current subtitle after successful card mining so newly known words recolor immediately.
- **Linux Multi-Line Copy:** Fixed multi-line subtitle copy on Linux so the overlay handles the follow-up digit selection locally, eliminating the timeout.
- **mpv Crash Notifications:** Stopped mpv from owning long-running SubMiner subprocesses during shutdown, preventing desktop crash notifications when closing video.
- **mpv Buffering Reload:** Kept the overlay alive across same-media reloads during buffering, avoiding duplicate startup gates and AniSkip lookups.
- **mpv Playlist Navigation:** Playlist navigation now reuses the running overlay without repeating the pause-until-ready warmup gate.
- **Anki Manual Updates:** Manual clipboard subtitle updates now replace sentence audio and media even when audio overwrite is disabled, while preserving existing word audio.
- **AniList Progress Tracking:** Fixed post-watch progress to check on mpv time-position updates using the fresh mpv position, wired manual mark-watched to force a progress sync, and filled missing episode metadata from the filename parser. Prevented duplicate post-watch writes during concurrent checks and preserved manual watched marks when sync fails.
- **AniList Setup:** Prevented config reload from opening the setup window during playback when token storage is unavailable; stopped setup from reporting success when encrypted token persistence fails.
- **AniList Linux Keyring:** Retried keyring availability after transient GNOME libsecret startup failures so stored tokens load once the keyring becomes available.
- **Stats Daemon:** Fixed stats background mode to route through the isolated stats daemon and deferred in-app stats startup when a daemon is already running, so video launches no longer fail when the stats port is in use.
- **Stats Session Detail:** Fixed recent session detail pages showing "Media not found" before lifetime media summaries are available.
- **Hover Background Config:** Accepted `subtitleStyle.hoverBackground` as an alias for `subtitleStyle.hoverTokenBackgroundColor` so setting it to `transparent` removes hover token backgrounds.
- **Managed Playback Exit:** Launcher-managed playback now exits the background SubMiner app when the video closes; explicit background launches remain persistent.
- **Multi-Mine Shortcuts:** Multi-line subtitle mining now accepts follow-up number-row digits even when the original shortcut modifiers are still held.
- **Jellyfin Setup:** Improved setup with recent server selection and inline authentication feedback; added a tray "Jellyfin Discovery" toggle for runtime-only cast discovery.
- **Subtitle POS Filtering:** Used Yomitan `wordClasses` metadata for subtitle part-of-speech filtering and backfilled blank MeCab POS detail fields during parser enrichment.

### Docs

- Improved the docs homepage with canonical URLs and a cleaner sitemap for better search engine indexing.

<details>
<summary>Internal changes</summary>

### Internal

- Replaced the changelog renderer with an AI polish pass that merges related fragments and writes user-facing release notes. `CHANGELOG.md` keeps internal items in a collapsed `<details>` block; GitHub release notes omit them entirely.
- Release CI no longer auto-builds pending `changes/*.md` fragments on tag. Tagging now fails fast if fragments remain — run `bun run changelog:build` (requires the `claude` CLI) and commit before tagging.

</details>

## v0.12.0 (2026-04-11)

### Changed

- Overlay: Added configurable overlay shortcuts for session help, controller select, and controller debug actions.
- Overlay: Added mpv/plugin and CLI routing for session help, controller utilities, and subtitle sidebar toggling through the shared session-action path.
- Overlay: Improved dedicated overlay modal retry and focus handling for runtime options, Jimaku, session help, controller tools, and the playlist browser.
- Overlay: Fixed controller configuration and controller debug shortcut opens so configured bindings bring up their modals again instead of tripping renderer recovery.
- Stats: Sessions are rolled up per episode within each day, with a bulk delete that wipes every session in the group.
- Stats: Trends add a 365-day range next to the existing 7d/30d/90d/all options.
- Stats: Library detail view gets a delete-episode action that removes the video and all its sessions.
- Stats: Vocabulary Top 50 tightens the word/reading column so katakana entries no longer push the scores off screen.
- Stats: Episode detail hides card events whose Anki notes have been deleted, instead of showing phantom mining activity.
- Stats: Trend and watch-time charts share a unified theme with horizontal gridlines and larger ticks for legibility.
- Stats: Overview, Library, Trends, Sessions, and Vocabulary now use generic "title" wording so YouTube videos and anime live comfortably side by side in the dashboard.
- Stats: Session timeline no longer plots seek-forward/seek-backward markers — they were too noisy on sessions with lots of rewinds.
- Stats: Replaced the "Library — Per Day" section on the Stats → Trends page with a "Library — Summary" section. The new section shows a top-10 watch-time leaderboard chart and a sortable per-title table (watch time, videos, sessions, cards, words, lookups, lookups/100w, date range), all scoped to the current date range selector.

### Fixed

- Overlay: Fixed overlay drag-and-drop routing so dropping external subtitle files like `.ass` onto mpv still loads them when the overlay is visible.
- Overlay: Addressed the latest CodeRabbit follow-ups on PR #49, including generation-scoped Lua session binding names, stricter session command validation, session-help shortcut visibility, the numeric-selection key guard, stats-overlay startup classification, and safer session-binding persistence.
- Overlay: Addressed the latest CodeRabbit follow-ups on the Windows overlay flow, including exact mpv target resolution, lower-overlay helper arguments, Win32 failure detection, and overlay cleanup on tracker loss.
- Overlay: Fixed Windows overlay z-order so the visible subtitle overlay stops staying above unrelated apps after mpv loses focus.
- Overlay: Fixed Windows overlay tracking to use native window polling and owner/z-order binding, which keeps the subtitle overlay aligned to the active mpv window more reliably.
- Overlay: Fixed Windows overlay hide/restore behavior so minimizing mpv immediately hides the overlay and restoring mpv brings it back on top of the mpv window without requiring a click.
- Overlay: Fixed stats overlay layering so the in-player stats page now stays above mpv and the subtitle overlay while it is open.
- Overlay: Fixed Windows subtitle overlay stability so transient tracker misses and restore events keep the current subtitle visible instead of waiting for the next subtitle line.
- Overlay: Fixed Windows focus handoff from the interactive subtitle overlay back to mpv so the overlay no longer drops behind mpv and briefly disappears.
- Overlay: Fixed Windows visible-overlay startup so it no longer briefly opens as an interactive or opaque surface before the tracked transparent overlay state settles.
- Overlay: Fixed spurious auto-pause after overlay visibility recovery and window resize so the overlay no longer pauses mpv until the pointer genuinely re-enters the subtitle area.
- Overlay: Fixed Windows secondary subtitle hover mode so the expanded hover hit area no longer blocks the native minimize, maximize, and close buttons.
- Overlay: Fixed Windows Yomitan popup focus loss after closing nested lookups so the original popup stays interactive instead of falling through to mpv.
- Stats: Fixed immersion-tracker timestamp handling under Bun/libsql so library rows, session timelines, and lifetime summaries keep real wall-clock millisecond values instead of truncating to invalid negative timestamps.
- Mpv Plugin: Fixed the mpv Lua plugin so hover and environment modules no longer use the `goto continue` pattern that can fail to parse on some user Lua runtimes.

### Internal

- Release: Added a dedicated beta/rc prerelease GitHub Actions workflow that publishes GitHub prereleases without consuming pending changelog fragments or updating AUR.
- Release: Added prerelease note generation so beta and release-candidate tags can reuse the current pending `changes/*.md` fragments while leaving stable changelog publication for the final release cut.

## v0.11.2 (2026-04-07)

### Changed

- Launcher: Replaced the launcher-only fullscreen toggle with `mpv.launchMode` so SubMiner-managed mpv playback can start in normal, maximized, or fullscreen mode.

### Fixed

- Launcher: Fixed launcher-managed mpv spawning to force an explicit X11 GPU path when Wayland trackers are unavailable.
- Launcher: Local playback now promotes a single unlabeled external subtitle sidecar to the primary slot instead of leaving mpv's embedded English auto-selection in place.
- Release: Fixed Linux AppImage startup packaging so Chromium child relaunches can resolve the bundled `libffmpeg.so` instead of crash-looping on startup.

## v0.11.1 (2026-04-04)

### Fixed

- Release: Linux packaged builds now expose the canonical `SubMiner` app identity to Electron's startup metadata so native Wayland compositors stop reporting the window class/app-id as lowercase `subminer`.
- Linux: Linux now restores the runtime options, Jimaku, and Subsync shortcuts after the Electron 39 regression by routing those actions through the overlay's mpv/IPC shortcut path.

## v0.11.0 (2026-04-03)

### Added

- Overlay: Added a playlist browser overlay modal for browsing sibling video files and the live mpv queue during playback.
- Overlay: Added the default `Ctrl+Alt+P` keybinding to open the playlist browser and manage queue order without leaving playback.

### Changed

- Setup: Made mpv plugin installation mandatory in the first-run setup flow, removed the skip path, and kept Finish disabled until the plugin is installed.
- Setup: Clarified that the mpv plugin requirement applies to setup on every platform, while the optional `SubMiner mpv` shortcut remains the recommended Windows playback entry point.
- Launcher: Streamlined Windows setup and config by making the `SubMiner mpv` shortcut self-contained and keeping `mpv.executablePath` as the simple fallback when `mpv.exe` is not on `PATH`.
- Overlay: Changed fresh-install default config to keep texthooker and stats from auto-opening browser tabs.
- Overlay: Changed fresh-install default config to enable AnkiConnect, Discord Rich Presence, subtitle-sidebar, and Yomitan-popup auto-pause by default, while disabling controller input by default.

### Fixed

- Main: Resolve the YouTube playback socket path lazily so startup honors CLI and config overrides.
- Main: Add regression coverage for the lazy socket-path lookup during Windows mpv startup.
- Main: Keep integrated `--start --texthooker` launches on the full app-ready startup path so the texthooker page and websocket servers start together during normal playback startup.
- Main: Stop the mpv/plugin auto-start flow from spawning a separate standalone texthooker helper during normal `subminer <video>` launches.
- Overlay: Keep tracked macOS visible overlays click-through by default so subtitle sidebar passthrough works immediately without requiring a subtitle hover cycle first.
- Overlay: Add regression coverage for the macOS visible-overlay passthrough default.
- Anilist: Stop AniList post-watch from sending a second progress update when the current episode was already satisfied by a ready retry item in the same watch-completion pass.
- Anilist: Add regression coverage for the retry-queue plus live-update duplicate path.
- Overlay: Fixed Kiku duplicate grouping to reuse duplicate note IDs from both generic sentence-card creation and Yomitan popup mining instead of running extra duplicate scans after add.
- Overlay: Fixed the Yomitan popup mining flow to add cards in the background while keeping the stock popup progress feedback, then pause playback and close the lookup popup before the Kiku merge modal opens.
- Overlay: Fixed configured subtitle-jump keybindings so backward and forward subtitle seeks keep playback paused when invoked from a paused state.
- Launcher: Fixed the Windows `SubMiner mpv` shortcut and `SubMiner.exe --launch-mpv` flow to launch mpv with SubMiner's required default args directly instead of requiring an `mpv.conf` profile named `subminer`.
- Launcher: Clarified the Windows install and usage docs so the shortcut path is documented as self-contained, while the optional `subminer` mpv profile remains available for manual mpv launches.
- Launcher: Hardened the first-run setup blocker copy and stale custom-scheme handling so setup messages stay aligned with config, plugin, and dictionary readiness.
- Launcher: Fixed the Windows `SubMiner mpv` shortcut idle launch so loading a video after opening the shortcut keeps mpv in the expected SubMiner-managed session, auto-starts the overlay, and re-arms subtitle auto-selection for the newly opened file.
- Launcher: Removed the redundant `.` subtitle search path from the Windows shortcut launch args and deduped repeated subtitle source tracks in the manual sync picker so duplicate external subtitle entries no longer appear from the shortcut path.
- Playback: Fixed managed local playback so duplicate startup-ready retries no longer unpause media after a later manual pause on the same file.
- Playback: Fixed managed local subtitle auto-selection so local files reuse configured primary and secondary subtitle language priorities instead of staying on mpv's initial `sid=auto` guess.
- Launcher: Added a blank-by-default `mpv.executablePath` override for Windows playback so users can point SubMiner at `mpv.exe` when it is not on `PATH`.
- Launcher: Kept the Windows shortcut and `--launch-mpv` flow simple by preserving PATH auto-discovery as the default and exposing the override in first-run setup.
- Launcher: Added `windows` as a recognized launcher backend option and auto-detection target on Windows.
- Launcher: Honored `SUBMINER_YTDLP_BIN` consistently across YouTube playback URL resolution, track probing, subtitle downloads, and metadata probing.
- Launcher: Kept the first-run setup window from navigating away on unexpected URLs.
- Launcher: Made Windows mpv honor an explicitly configured executable path instead of silently falling back to PATH.
- Launcher: Hardened `--launch-mpv` parsing and Windows binary resolution so valueless flags do not swallow media targets and symlinked launcher installs do not short-circuit PATH lookup.
- Launcher: Fixed first-run setup blocking playback on macOS when the SubMiner mpv plugin was already installed at the canonical `~/.config/mpv` path.
- Launcher: Fixed setup gating so stale cancelled setup state no longer prevents playback when the canonical mpv plugin entrypoint already exists.
- Playback: Prevented stale async playlist-browser subtitle rearm callbacks from overriding newer subtitle selections during rapid file changes.

### Docs

- Docs Site: Added a dedicated Subtitle Sidebar guide and linked it from the homepage and configuration docs.
- Docs Site: Linked Jimaku integration from the homepage to its dedicated docs page.
- Docs Site: Refreshed docs-site theme tokens and hover/selection styling for the updated pages.

### Internal

- Release: Retried AUR clone and push operations in the tagged release workflow.
- Release: Kept GitHub Releases green when AUR publish flakes and needs manual follow-up.
- Release: Updated Electron to 39.8.6 and pinned patched transitive build dependencies to clear the reported high-severity audit findings.

## v0.10.0 (2026-03-29)

### Changed

- Integrations: Replaced the deprecated Discord Rich Presence wrapper with the maintained `@xhayper/discord-rpc` package.

### Fixed

- Stats: Fixed stats startup so the immersion tracker can run when `Bun.serve` is unavailable.
- Stats: Stats server now falls back to a Node `http` listener in Electron/runtime paths that do not expose Bun.
- Overlay: Fixed the macOS visible-overlay toggle path so manual hides stay hidden and the plugin uses the explicit visible-overlay toggle command.
- Subtitle Sidebar: Restored macOS mpv passthrough while the overlay subtitle sidebar is open so clicks outside the sidebar can refocus mpv and keep native keybindings working.

### Internal

- Release: Added a maintained source coverage lane that shards Bun coverage one test file at a time and merges LCOV output into `coverage/test-src/lcov.info`.
- Release: CI and release quality-gate now upload the merged source-lane LCOV artifact for inspection.
- Runtime: Extracted remaining inline runtime logic from `src/main.ts` into dedicated runtime modules and composer helpers.
- Runtime: Added focused regression tests for the extracted runtime/composer boundaries.
- Runtime: Updated task tracking notes to mark TASK-238.6 complete and confirm follow-on boot-phase split can be deferred.
- Runtime: Split `src/main.ts` boot wiring into dedicated `src/main/boot/services.ts`, `src/main/boot/runtimes.ts`, and `src/main/boot/handlers.ts` modules.
- Runtime: Added focused tests for the new boot-phase seams and kept the startup/typecheck/build verification lanes green.
- Runtime: Updated internal architecture/task docs to record the boot-phase split and new ownership boundary.

## v0.9.3 (2026-03-25)

### Changed

- Launcher: Moved YouTube primary subtitle language defaults to `youtube.primarySubLanguages`.
- Launcher: Removed the placeholder YouTube subtitle retime step and now uses downloaded primary subtitle tracks directly, so there is no fake path rewrite before playback/sidebar loading.
- YouTube: Removed the `src/core/services/youtube/retime` helper and its tests after retiring the internal retime strategy.
- Docs: Clarified optional `alass` / `ffsubsync` subtitle-sync requirements and setup steps, including fallback behavior when sync tools are absent.
- Launcher: Removed the old `youtubeSubgen.primarySubLanguages` config path from the generated config and docs.

## v0.9.2 (2026-03-25)

### Fixed

- Overlay: Fixed overlay pointer tracking so Windows click-through toggles immediately when the cursor enters or leaves subtitle regions, without waiting for a later hover resync.
- Overlay: Fixed Windows overlay window tracking on scaled displays by converting native tracked window bounds to Electron DIP coordinates before applying overlay bounds.
- Launcher: Fixed Windows direct `--youtube-play` startup so MPV boots reliably, stays paused until the app-owned subtitle flow is ready, and reuses an already-running SubMiner instance when available.
- Launcher: Fixed standalone Windows `--youtube-play` sessions so closing MPV fully exits SubMiner instead of leaving hidden overlay windows or a background process behind.
- Overlay: Fixed `subminer <youtube-url>` on Linux so the YouTube playback flow waits for Yomitan to load before creating the overlay window, avoiding the broken lookup popup state that previously required a manual overlay refresh.

## v0.9.1 (2026-03-24)

### Changed

- Release: Reduced packaged release size by excluding duplicate `extraResources` payload and pruning docs, tests, sourcemaps, and other source-only files from Electron bundles.

### Fixed

- Overlay: Restored controller navigation and lookup/mining controls while the subtitle sidebar is open, while keeping true modal dialogs blocking controller actions.
- Tokenizer: Fixed subtitle annotation clearing so explanatory contrast endings like `んですけど` are excluded consistently across the shared tokenizer filter and annotation stage.

## v0.9.0 (2026-03-23)

### Added

- Docs: Added a new WebSocket / Texthooker API and integration guide covering WebSocket payloads, custom client patterns, mpv plugin automation, and webhook-style relay examples. Linked from configuration and mining workflow docs for easier discovery.

### Changed

- Launcher: Added an app-owned YouTube subtitle flow that pauses mpv, uses absPlayer-style YouTube timedtext parsing/conversion to download subtitle tracks, and injects them as external files before playback resumes.
- Launcher: Changed YouTube subtitle startup to auto-load the best-available primary and secondary subtitle tracks at launch instead of forcing the picker modal first. Secondary subtitle failures no longer block playback resume.
- Launcher: Added `Ctrl+Alt+C` as the default keybinding to manually open the YouTube subtitle picker during active YouTube playback.
- Launcher: Added yt-dlp metadata probing so YouTube playback and immersion tracking record canonical video title and channel metadata.
- Launcher: Stopped forcing `--ytdl-raw-options=` before user-provided mpv options so existing YouTube cookie integrations in user `--args` are no longer clobbered.
- Launcher: Disabled mpv native YouTube subtitle auto-loading for the app-owned flow so injected external subtitle files remain authoritative.
- Launcher: Added OSD status messages for YouTube playback startup, subtitle acquisition, and subtitle loading so the flow stays visible before and during the picker.
- Subtitle Sidebar: Added startup-auto-open controls and resume positioning improvements so the sidebar jumps directly to the first resolved active cue.
- Subtitle Sidebar: Improved subtitle prefetch and embedded overlay passthrough sync so sidebar and overlay subtitle states stay consistent across media transitions.
- Subtitle Sidebar: Updated scroll handling, embedded layout styling, and active-cue visual behavior.
- Stats: Stats Library tab now displays YouTube video title, channel name, and channel thumbnail for YouTube media entries, with retry logic to fill in metadata that arrives after initial load.

### Fixed

- Launcher: Fixed Anki media mining for mpv YouTube streams by unwrapping the stream URL so audio and screenshot capture work correctly for YouTube playback sessions.
- Immersion: Fixed YouTube media path handling in the immersion runtime and tracking so YouTube sessions record correct media references, AniList guessing skips YouTube URLs, and post-watch state transitions do not fire for YouTube media.
- Launcher: Fixed startup-launched YouTube playback so primary subtitle overlay updates continue after auto-load completes.
- Launcher: Fixed auto-loaded YouTube primary subtitles so parsed cues appear in the subtitle sidebar without needing a manual picker retry.
- Launcher: Fixed the YouTube picker to guard against duplicate subtitle submissions and tightened YouTube URL detection so follow-up runtime flows only treat real YouTube hosts as YouTube playback.
- Launcher: Fixed primary subtitle failure notifications being shown while app-owned YouTube subtitle probing and downloads are still in flight.
- Launcher: Preserved existing authoritative YouTube subtitle tracks when available; downloaded tracks are used only to fill missing sides, and native mpv secondary subtitle rendering is hidden so the overlay remains the sole secondary display.

## v0.8.0 (2026-03-22)

### Added

- Overlay: Added the subtitle sidebar feature with a new `subtitleSidebar` configuration surface and rendered sidebar modal with cue list rendering, click-to-seek, active-cue highlighting, and embedded layout support.
- IPC: Added sidebar snapshot plumbing between renderer and main process for overlay/sidebar synchronization.

### Changed

- Config: Added hot-reloadable sidebar options for enablement, layout, visibility, typography, opacity, sizing, and interaction behavior (`autoOpen`, `pauseOnHover`, `autoScroll`, toggle key).
- Docs: Added full `subtitleSidebar` documentation coverage, including sample config, option table, and toggle shortcut notes.
- Runtime: Improved subtitle prefetch/rendering flow so sidebar and overlay subtitle states stay in sync across media transitions.

### Fixed

- Overlay: Kept sidebar cue tracking stable across playback transitions and timing edge cases.
- Overlay: Improved sidebar resume/start behavior to jump directly to the first resolved active cue.
- Overlay: Stopped stale subtitle refreshes from regressing active-cue and text state.

## v0.7.0 (2026-03-19)

### Added

- Immersion: Added Mine Word, Mine Sentence, and Mine Audio buttons to word detail example lines in the stats dashboard.
- Immersion: Mine Word creates a full Yomitan card (definition, reading, pitch accent) via the hidden search page bridge, then enriches with sentence audio, screenshot, and metadata extracted from the source video.
- Immersion: Mine Sentence and Mine Audio create cards directly with appropriate Lapis/Kiku flags, sentence highlighting, and media from the source file.
- Immersion: Media generation (audio + image/AVIF) runs in parallel and respects all AnkiConnect config options.
- Immersion: Added word exclusion list to the Vocabulary tab with localStorage persistence and a management modal.
- Immersion: Fixed truncated readings in the frequency rank table (e.g. お前 now shows おまえ instead of まえ).
- Immersion: Clicking a bar in the Top Repeated Words chart now opens the word detail panel.
- Immersion: Secondary subtitle text is now stored alongside primary subtitle lines for use as translation when mining cards from the stats page.
- Stats: Added `subminer stats -b` to start or reuse a dedicated background stats server without blocking normal SubMiner instances.
- Stats: Added `subminer stats -s` to stop the dedicated background stats server without closing browser tabs.
- Stats: Stats server startup now reuses a running background stats daemon instead of trying to bind a second local server in another SubMiner instance.
- Launcher: Added launcher passthrough for `-a/--args` so mpv receives raw extra launch flags (`--fs`, `--ytdl-format`, custom audio/video settings, etc.) from the `subminer` command.
- Launcher: Added `subminer stats` to launch the local stats dashboard, force-start the stats server on demand, and open the dashboard in your browser.
- Launcher: Added `subminer stats cleanup` to backfill vocabulary metadata and prune stale or excluded immersion rows on demand.
- Launcher: Added `stats.autoOpenBrowser` so browser launch after `subminer stats` can be enabled or disabled explicitly.
- Immersion: Added a local stats dashboard for immersion tracking with Overview, Anime, Trends, Vocabulary, and Sessions views.
- Immersion: Added anime progress, episode completion, Anki card links, and occurrence drill-down across the stats dashboard.
- Immersion: Added richer session timelines with new-word activity, cumulative totals, and pause/seek/card event markers.
- Immersion: Added completed-episodes and completed-anime totals to the Overview tracking snapshot.

### Changed

- Anki: Changed known-word cache settings to live under `ankiConnect.knownWords` instead of mixing them into `ankiConnect.nPlusOne`.
- Anki: Kept legacy `ankiConnect.nPlusOne` known-word keys and older `ankiConnect.behavior.nPlusOne*` keys as deprecated compatibility fallbacks.
- Stats: Added session deletion to the Sessions tab with the same confirmation prompt used by anime episode/session deletes, and removed all associated session rows from the stats database.
- Immersion: Kept immersion tracking history by default while preserving daily/monthly rollup maintenance.
- Immersion: Added exact lifetime summary reads for overview/anime/media stats so dashboard totals no longer depend on rescanning raw telemetry.
- Immersion: Reduced tracker storage overhead by removing duplicated subtitle text from subtitle-line event payloads.
- Immersion: Deduplicated episode cover-art blobs through a shared blob store and updated cover-art reads/writes to resolve shared images correctly.
- Immersion: Added indexes for large-history session, telemetry, vocabulary, kanji, and cover-art queries to keep dashboard reads fast as the SQLite database grows.
- Immersion: Renamed the stats dashboard's Anime tab to Library so the media browser label matches non-anime sources like YouTube and other yt-dlp-backed content.
- Anilist: Standardized episode completion threshold by introducing `DEFAULT_MIN_WATCH_RATIO` and using it for both local watched state transitions and AniList post-watch progress updates.
- Anilist: Episode auto-marking now uses the same threshold as AniList (`85%`), removing divergent completion behavior.
- Overlay: Excluded interjections and sound-effect tokens from subtitle annotation styling so they no longer inherit misleading lexical highlight treatment while still remaining visible and hoverable as plain subtitle tokens.
- Overlay: Expanded subtitle annotation noise filtering to also strip annotation metadata from standalone grammar-only helper tokens such as particles, auxiliaries, adnominals, common explanatory endings like `んです` / `のだ`, and merged trailing quote-particle forms like `...って` while keeping them tokenized for hover lookup.

### Fixed

- Launcher: Fixed mpv Lua plugin binary auto-detection on Linux to also search `/usr/bin/subminer` and `/usr/local/bin/subminer` (lowercase), matching the conventional Unix wrapper name used by packaged installs such as the AUR package.
- Stats: Fixed the in-app stats overlay so it connects to the configured `stats.serverPort` instead of falling back to the default port.
- Overlay: Fixed subtitle frequency tagging for merged lookup-backed tokens like `陰に` by falling back to exact surface-form Yomitan frequencies when the normalized headword lookup misses.
- Overlay: Fixed MeCab merged-token position mapping across line breaks so merged content-plus-particle tokens like `陰に` keep their matched Yomitan frequency instead of inheriting shifted POS tags.
- Overlay: Fixed grouped frequency parsing in both Yomitan and fallback frequency-dictionary lookups so display values like `118,121` use the leading rank instead of collapsing the rank and occurrence count into `118121`.
- Overlay: Fixed frequency-rank ingestion to ignore Yomitan dictionaries explicitly marked `occurrence-based`, so raw occurrence counts are no longer treated as subtitle rank values.
- Overlay: Fixed inflected headword frequency tagging to prefer ranks from the selected Yomitan `termsFind` popup entry itself, ordered by configured dictionary priority, so forms like `潜み` use primary-dictionary ranks like `4073` before falling back to lower-priority raw lemma metadata such as `CC100`.
- Overlay: Fixed annotation-stage frequency filtering so exact kanji noun tokens like `者` keep their matched rank even when MeCab labels them `名詞/非自立`, instead of dropping the highlight after scan-time frequency lookup succeeds.
- Anki: Fixed repeated character-dictionary startup work by scheduling auto-sync only from mpv media-path changes instead of also re-triggering it from connection and media-title events for the same title.
- Overlay: Fixed macOS fullscreen overlay stability by keeping the passive visible overlay from stealing focus, re-raising the overlay window when reasserting its macOS topmost level, and tolerating one transient macOS tracker/helper miss before hiding the overlay.
- Overlay: Kept subtitle tokenization warmup one-shot for the lifetime of the app so later fullscreen/media churn on macOS does not replay the startup warmup gate after the first file is ready.
- Overlay: Added a bounded macOS tracker loss-grace window so fullscreen enter/leave transitions do not immediately hide and reload the overlay when the helper briefly loses the mpv window.
- Overlay: Skipped subtitle/tokenization refresh invalidation on character-dictionary auto-sync completion when the dictionary was already current, preventing startup flash/reload loops on unchanged media.
- Stats: Fixed session stats so known-word counts track real known-word occurrences without collapsing subtitle-line gaps.
- Stats: Fixed session word totals in session-facing stats views to prefer token counts when available, preventing known words from exceeding total words in the session chart.
- Stats: Fixed the stats Vocabulary tab blank-screen regression caused by a hook-order crash after vocabulary data finished loading.
- Anki: Fixed card-mine OSD feedback so the final mine result stops the Anki spinner first, then shows a single-line `✓`/`x` status without being overwritten by a later spinner tick.
- Stats: Removed the misleading `New words` series from expanded session charts; session detail now shows only the real total-word and known-word lines.
- Stats: Restored the cross-anime word table behavior in stats vocabulary surfaces so shared vocabulary entries no longer disappear or merge incorrectly across related media.
- Stats: `subminer stats -b` now runs as a standalone background stats daemon instead of reusing the main SubMiner app process, so the overlay app can still be launched separately for normal video watching.
- Stats: Dashboard word mining still works against the background daemon by using a short-lived hidden helper for the Yomitan add-note flow.
- Stats: Load full session timelines by default in stats session detail views so long sessions preserve complete telemetry history instead of being truncated by a fixed sample limit.
- Stats: Replaced heuristic stats word counts with Yomitan token counts, so session, media, anime, and trend subtitle totals now come directly from parsed subtitle tokens.
- Stats: Updated stats UI labels and lookup-rate copy to refer to tokens instead of words where those counts are shown.
- Overlay: Reduced repeated `Overlay loading...` popups on macOS when fullscreen tracker flaps briefly hide and recover the visible overlay.
- Stats: Scaled expanded session-detail known-word charts to the session's actual percentage range so small changes no longer render as a nearly flat line.
- Jlpt: Reduced JLPT dictionary startup log noise by summarizing duplicate surface-form collisions instead of logging one line per duplicate entry.

## v0.6.5 (2026-03-15)

### Internal

- Release: Seed the AUR checkout with the repo `.SRCINFO` template before rewriting metadata so tagged releases do not depend on prior AUR state.

## v0.6.4 (2026-03-15)

### Internal

- Release: Reworked AUR metadata generation to update `.SRCINFO` directly instead of depending on runner `makepkg`, fixing tagged release publishing for `subminer-bin`.

## v0.6.3 (2026-03-15)

### Changed

- Overlay: Expanded the `Alt+C` controller modal into an inline config/remap flow with preferred-controller saving and per-action learn mode for buttons, triggers, and stick directions.

### Internal

- Workflow: Hardened the `subminer-scrum-master` skill to explicitly answer whether docs updates and changelog fragments are required before handoff.
- Release: Automate `subminer-bin` AUR package updates from the tagged release workflow.

## v0.6.2 (2026-03-12)

### Changed

- Config: Added `yomitan.externalProfilePath` to reuse another Electron app's Yomitan profile in read-only mode.
- Config: SubMiner now reuses external Yomitan dictionaries/settings without writing back to that profile.
- Config: Launcher-managed playback now respects `yomitan.externalProfilePath` and no longer forces first-run setup when external Yomitan is configured.
- Config: SubMiner now seeds `config.jsonc` even when the default config directory already exists.
- Config: First-run setup now allows zero internal dictionaries when `yomitan.externalProfilePath` is configured, and falls back to requiring at least one internal dictionary if that external profile is later removed.

## v0.6.1 (2026-03-12)

### Added

- Overlay: Added Chrome Gamepad API controller support for keyboard-only overlay mode, including configurable logical bindings for lookup, mining, popup navigation, Yomitan audio, mpv pause, d-pad fallback navigation, and slower smooth popup scrolling.
- Overlay: Added `Alt+C` controller selection and `Alt+Shift+C` controller debug modals, with preferred controller persistence and live raw input inspection.
- Overlay: Added a transient in-overlay controller-detected indicator when a controller is first found.
- Overlay: Fixed stale keyboard-only token highlight cleanup when keyboard-only mode turns off or the Yomitan popup closes.

### Docs

- Install: Added Arch Linux AUR install docs for `subminer-bin` in the README and installation guide.

### Internal

- Config: add an enforced `verify:config-example` gate so checked-in example config artifacts cannot drift silently
- Release: Fixed the release workflow token permissions so tagged builds can download `oven-sh/setup-bun` and publish artifacts again.

## v0.5.6 (2026-03-10)

### Fixed

- Dictionary: Persist merged character-dictionary MRU state as soon as a new retained set is built so revisits do not get dropped if later Yomitan import work fails, and skip merged dictionary rebuilds for reorder-only revisits when the retained anime set itself has not changed.
- Startup: Fixed early Electron startup writing config and user data under a lowercase `~/.config/subminer` path instead of the canonical `~/.config/SubMiner` directory.
- Overlay: Kept JLPT underline colors stable during Yomitan hover and selection states, even when tokens also use known, N+1, name-match, or frequency styling.

## v0.5.5 (2026-03-09)

### Changed

- Overlay: Added `f` as the default overlay fullscreen toggle and changed the default AniSkip intro-jump key to `Tab`.
- Dictionary: Aligned AniList character dictionary generation more closely with the upstream reference by preserving duplicate shared names across characters, skipping characters without native Japanese names, restoring richer character info fields, and using upstream-style role mapping plus hint-aware kanji readings.
- Startup: Ordered startup OSD messages so tokenization loads first, annotation loading appears next if still pending, and character dictionary sync progress waits until annotation loading finishes.
- Dictionary: Added a visible startup OSD step for merged character-dictionary building so long rebuilds show progress before the later import/upload phase.

### Fixed

- Dictionary: Fixed AniList media guessing for character dictionary auto-sync by using filename-only `guessit` input and preserving multi-part guessit titles instead of truncating them to the first segment.
- Dictionary: Refresh the current subtitle after character dictionary auto-sync completes so newly imported character names highlight on the active line instead of waiting for the next subtitle change.
- Dictionary: Show character dictionary auto-sync progress on the mpv OSD without sending desktop notifications.
- Dictionary: Keep character dictionary auto-sync non-blocking during startup by letting snapshot/build work run in parallel and delaying only the Yomitan import/settings phase until current-media tokenization is already ready.
- Overlay: Fixed visible overlay keyboard handling so pressing `Tab` still reaches mpv and triggers the default AniSkip skip-intro binding while the overlay has focus.
- Plugin: Fix Windows mpv plugin binary override lookup so `SUBMINER_BINARY_PATH` still resolves to `SubMiner.exe` when no AppImage override is set.

## v0.5.3 (2026-03-09)

### Changed

- Release: Publish unsigned Windows `.exe` and `.zip` artifacts directly from release CI instead of routing them through SignPath.
- Release: Added `bun run build:win:unsigned` for explicit local unsigned Windows packaging.

## v0.5.2 (2026-03-09)

### Internal

- Release: Pinned the Windows SignPath submission workflow to an explicit artifact-configuration slug instead of relying on the SignPath project's default configuration.

## v0.5.1 (2026-03-09)

### Changed

- Launcher: Removed the YouTube subtitle generation mode switch so YouTube playback always preloads subtitles before mpv starts.

### Fixed

- Launcher: Hardened YouTube AI subtitle fixing so fenced SRT output and text-only one-cue-per-block responses can still be applied without losing original cue timing.
- Launcher: Skipped AniSkip lookup during URL playback and YouTube subtitle-preload playback, limiting AniSkip to local file targets where it can actually resolve anime metadata.
- Launcher: Keep the background SubMiner process running after a launcher-managed mpv session exits so the next mpv instance can reconnect without restarting the app.
- Launcher: Reuse prior tokenization readiness after the background app is already warm so reopening a video does not pause again waiting for duplicate warmup completion.
- Windows: Acquire the app single-instance lock earlier so Windows overlay/video launches reuse the running background SubMiner process instead of booting a second full app and repeating startup warmups.

## v0.3.0 (2026-03-05)

- Added keyboard-driven Yomitan navigation and popup controls, including optional auto-pause.
- Added subtitle/jump keyboard handling fixes for smoother subtitle playback control.
- Improved Anki/Yomitan reliability with stronger Yomitan proxy syncing and safer extension refresh logic.
- Added Subsync `replace` option and deterministic retime naming for subtitle workflows.
- Moved aniskip resolution to launcher-script options for better control.
- Tuned tokenizer frequency highlighting filters for improved term visibility.
- Added release build quality-of-life for CLI publish (`gh`-based clobber upload).
- Removed docs Plausible integration and cleaned associated tracker settings.

## v0.2.3 (2026-03-02)

- Added performance and tokenization optimizations (faster warmup, persistent MeCab usage, reduced enrichment lookups).
- Added subtitle controls for no-jump delay shifts.
- Improved subtitle highlight logic with priority and reliability fixes.
- Fixed plugin loading behavior to keep OSD visible during startup.
- Fixed Jellyfin remote resume behavior and improved autoplay/tokenization interaction.
- Updated startup flow to load dictionaries asynchronously and unblock first tokenization sooner.

## v0.2.2 (2026-03-01)

- Improved subtitle highlighting reliability for frequency modes.
- Fixed Jellyfin misc info formatting cleanup.
- Version bump maintenance for 0.2.2.

## v0.2.1 (2026-03-01)

- Delivered Jellyfin and Subsync fixes from release patch cycle.
- Version bump maintenance for 0.2.1.

## v0.2.0 (2026-03-01)

- Added task-related release work for the overlay 2.0 cycle.
- Introduced Overlay 2.0.
- Improved release automation reliability.

## v0.1.2 (2026-02-24)

- Added encrypted AniList token handling and default GNOME keyring support.
- Added launcher passthrough for password-store flows (Jellyfin path).
- Updated docs for auth and integration behavior.
- Version bump maintenance for 0.1.2.

## v0.1.1 (2026-02-23)

- Fixed overlay modal focus handling (`grab input`) behavior.
- Version bump maintenance for 0.1.1.

## v0.1.0 (2026-02-23)

- Bootstrapped Electron runtime, services, and composition model.
- Added runtime asset packaging and dependency vendoring.
- Added project docs baseline, setup guides, architecture notes, and submodule/runtime assets.
- Added CI release job dependency ordering fixes before launcher build.
