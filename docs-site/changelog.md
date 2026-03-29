# Changelog

## v0.10.0 (2026-03-29)
- Fixed stats startup so the immersion tracker can run when `Bun.serve` is unavailable.
- Added a Node `http` fallback for Electron/runtime paths that do not expose Bun, so stats keeps working there too.
- Updated Discord Rich Presence to the maintained `@xhayper/discord-rpc` wrapper.
- Fixed the macOS visible-overlay toggle path so manual hides stay hidden and the plugin uses the explicit visible-overlay toggle command.
- Restored macOS mpv passthrough while the overlay subtitle sidebar is open so clicks outside the sidebar can refocus mpv and keep native keybindings working.

## v0.9.3 (2026-03-25)
- Moved YouTube primary subtitle language defaults to `youtube.primarySubLanguages`.
- Removed the placeholder YouTube subtitle retime step; downloaded primary subtitle tracks are now used directly.
- Removed the old internal YouTube retime helper and its tests.
- Clarified optional `alass` / `ffsubsync` subtitle-sync setup and fallback behavior in the docs.
- Removed the legacy `youtubeSubgen.primarySubLanguages` config path from generated config and docs.

## v0.9.2 (2026-03-25)
- Fixed overlay pointer tracking so Windows click-through toggles immediately when the cursor enters or leaves subtitle regions.
- Fixed Windows overlay window tracking on scaled displays by converting native tracked window bounds to Electron DIP coordinates.
- Fixed Windows direct `--youtube-play` startup so MPV boots reliably, stays paused until the app-owned subtitle flow is ready, and reuses an already-running SubMiner instance.
- Fixed standalone Windows `--youtube-play` sessions so closing MPV fully exits SubMiner instead of leaving hidden overlay windows behind.
- Fixed `subminer <youtube-url>` on Linux so the YouTube playback flow waits for Yomitan to load before creating the overlay window.

## v0.9.1 (2026-03-24)
- Reduced packaged release size by excluding duplicate `extraResources` payload and pruning docs, tests, sourcemaps, and other source-only files from Electron bundles.
- Restored controller navigation and lookup/mining controls while the subtitle sidebar is open, while keeping true modal dialogs blocking controller actions.
- Fixed subtitle annotation clearing so explanatory contrast endings like `んですけど` are excluded consistently across the shared tokenizer filter and annotation stage.

## v0.9.0 (2026-03-23)
- Added an app-owned YouTube subtitle flow with absPlayer-style timedtext parsing that auto-loads the default primary subtitle plus a best-effort secondary at startup and resumes once the primary is ready.
- Added a manual YouTube subtitle picker on `Ctrl+Alt+C` so subtitle selection can be retried on demand during active YouTube playback.
- Added yt-dlp metadata probing so YouTube playback and immersion tracking record canonical video title and channel metadata.
- Disabled conflicting mpv native subtitle auto-selection for the app-owned flow so injected explicit tracks stay authoritative.
- Added OSD status updates covering YouTube playback startup, subtitle acquisition, and subtitle loading.
- Stopped forcing `--ytdl-raw-options=` before user-provided mpv options so existing YouTube cookie integrations are preserved.
- Improved sidebar startup/resume behavior, scroll handling, and overlay/sidebar subtitle synchronization.
- Stats Library tab now shows YouTube video title, channel name, and thumbnail for YouTube media entries.
- Added a new WebSocket / Texthooker API integration guide covering payload formats, custom client patterns, and mpv plugin automation.
- Fixed Anki media mining for mpv YouTube streams so audio and screenshot capture work correctly during YouTube playback sessions.
- Fixed YouTube media path handling in immersion tracking so YouTube sessions record correct media references and AniList state transitions do not fire for YouTube media.
- Reused existing authoritative YouTube subtitle tracks when present, fell back only for missing sides, and kept native mpv secondary subtitle rendering hidden so the overlay remains the visible secondary subtitle surface.

## v0.8.0 (2026-03-22)
- Added a configurable subtitle sidebar feature (`subtitleSidebar`) with overlay/embedded rendering, click-to-seek cue list, and hot-reloadable visibility and behavior controls.
- Added a rendered sidebar modal with cue list display, click-to-seek, active-cue highlighting, and embedded layout support.
- Added sidebar snapshot plumbing between main and renderer for overlay/sidebar synchronization.
- Added sidebar configuration options for visibility and behavior (enabled, layout, toggle key, autoOpen, pauseOnHover, autoScroll) plus typography and sizing controls.
- Documented `subtitleSidebar` configuration and behavior in user-facing docs (configuration.md, shortcuts.md, config.example.jsonc).
- Updated subtitle prefetch/rendering flow to keep overlay and sidebar state in sync through media transitions.
- Kept sidebar cue tracking stable across playback transitions and timing edge cases.
- Fixed sidebar startup/resume positioning to jump directly to the first resolved active cue.
- Prevented stale subtitle refreshes from regressing active-cue state.

## v0.7.0 (2026-03-19)
- Added a full local immersion dashboard release line with Overview, Library, Trends, Vocabulary, and Sessions drill-down views backed by SQLite tracking data.
- Added browser-first stats workflows: `subminer stats`, background stats daemon controls (`-b` / `-s`), stats cleanup, and dashboard-side mining actions with media enrichment.
- Improved stats accuracy and scale handling with Yomitan token counts, full session timelines, known-word timeline fixes, cross-media vocabulary fixes, and clearer session charts.
- Improved overlay/runtime stability with quieter macOS fullscreen recovery, reduced repeated loading OSD popups, and better frequency/noise handling for subtitle annotations.
- Added launcher mpv-args passthrough plus Linux plugin wrapper-name fallback for packaged installs.
- Added a hover-revealed ↗ button on Sessions tab rows to navigate directly to the anime media-detail view, with correct "Back to Sessions" back-navigation.
- Excluded auxiliary-stem `そうだ` grammar tails (MeCab POS3 `助動詞語幹`) from subtitle annotation metadata so frequency, JLPT, and N+1 styling no longer bleed onto grammar-tail tokens.

## v0.6.5 (2026-03-15)
- Seeded the AUR checkout with the repo `.SRCINFO` template before rewriting metadata so tagged releases do not depend on prior AUR state.

## v0.6.4 (2026-03-15)
- Reworked AUR metadata generation to update `.SRCINFO` directly instead of depending on runner `makepkg`, fixing tagged release publishing for `subminer-bin`.

## v0.6.3 (2026-03-15)
- Expanded `Alt+C` into an inline controller config/remap flow with preferred-controller saving and per-action learn mode for buttons, triggers, and stick directions.
- Automated `subminer-bin` AUR package updates from the tagged release workflow.

## v0.6.2 (2026-03-12)
- Added `yomitan.externalProfilePath` so SubMiner can reuse another Electron app's Yomitan profile in read-only mode.
- Reused external Yomitan dictionaries/settings without writing back to that profile.
- Let launcher-managed playback honor external Yomitan config instead of forcing first-run setup.
- Seeded `config.jsonc` even when the default config directory already exists.
- Let first-run setup complete without internal dictionaries while external Yomitan is configured, then require an internal dictionary again only if that external profile is later removed.

## v0.6.1 (2026-03-12)
- Added Chrome Gamepad API controller support for keyboard-only overlay mode.
- Added configurable controller bindings for lookup, mining, popup navigation, Yomitan audio, mpv pause, and d-pad fallback navigation.
- Added smooth, slower popup scrolling for controller navigation.
- Expanded `Alt+C` into a controller config/remap modal with preferred-controller saving, inline learn mode, and kept `Alt+Shift+C` for raw input debugging.
- Added a transient in-overlay controller-detected indicator when a controller is first found.
- Fixed cleanup of stale keyboard-only token highlights when keyboard-only mode is disabled or when the Yomitan popup closes.
- Added an enforced `verify:config-example` gate so checked-in example config artifacts cannot drift silently.

## v0.5.6 (2026-03-10)
- Persisted merged character-dictionary MRU state as soon as a new retained set is built so revisits do not get dropped if later Yomitan import work fails.
- Fixed early Electron startup writing config and user data under a lowercase `~/.config/subminer` path instead of canonical `~/.config/SubMiner`.
- Kept JLPT underline colors stable during Yomitan hover and selection states, even when tokens also use known, N+1, name-match, or frequency styling.

## v0.5.1 (2026-03-09)
- Removed the old YouTube subtitle-generation mode switch; YouTube playback now resolves subtitles before mpv starts.
- Hardened YouTube AI subtitle fixing so fenced/text-only responses keep original cue timing.
- Skipped AniSkip during URL/YouTube playback where anime metadata cannot be resolved reliably.
- Kept the background SubMiner process warm across launcher-managed mpv exits so reconnects do not repeat startup pause/warmup work.
- Fixed Windows single-instance reuse so overlay and video launches reuse the running background app instead of booting a second full app.
- Hardened the Windows signing/release workflow with SignPath retry handling for signed `.exe` and `.zip` artifacts.

## v0.5.0 (2026-03-08)
- Added the initial packaged Windows release.
- Added Windows-native mpv window tracking, launcher/runtime plumbing, and packaged helper assets.
- Improved close behavior so ending playback hides the visible overlay while the background app stays running.
- Limited the native overlay outline/debug frame to debug mode on Windows.

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
