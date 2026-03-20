# Changelog

## v0.7.0 (2026-03-19)
- Added a full local immersion dashboard release line with Overview, Library, Trends, Vocabulary, and Sessions drill-down views backed by SQLite tracking data.
- Added browser-first stats workflows: `subminer stats`, background stats daemon controls (`-b` / `-s`), stats cleanup, and dashboard-side mining actions with media enrichment.
- Improved stats accuracy and scale handling with Yomitan token counts, full session timelines, known-word timeline fixes, cross-media vocabulary fixes, and clearer session charts.
- Improved overlay/runtime stability with quieter macOS fullscreen recovery, reduced repeated loading OSD popups, and better frequency/noise handling for subtitle annotations.
- Added launcher mpv-args passthrough plus Linux plugin wrapper-name fallback for packaged installs.

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
