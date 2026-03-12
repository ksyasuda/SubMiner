# Changelog

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
