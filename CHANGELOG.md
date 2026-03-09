# Changelog

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
