# Changelog

## Unreleased

### Fixed
- AniList: Stopped post-watch tracking from sending a second progress update when the current episode was already satisfied by a ready retry item in the same watch-completion pass.

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
