## Highlights
### Added

- **Character Dictionary:** Added AniList-based selection to resolve character dictionary mismatches, with series-scoped overrides that replace stale entries. Available via `subminer dictionary --candidates` / `--select` and a default `Ctrl+Alt+A` in-app shortcut.
- **Subtitle Bar Toggle:** Added a `V` shortcut and mpv binding to toggle the primary subtitle bar independently of mpv's native subtitle display.
- **Texthooker:** Added `subminer texthooker -o` and a tray menu item to open the local texthooker page in the default browser.

### Changed

- **mpv Plugin Setup:** Managed launches now inject the bundled plugin automatically. The setup flow can trash detected legacy global plugin files before launch, and legacy global install entrypoints have been removed so regular mpv playback is unaffected.
- **Tray Menu:** Replaced "Open Overlay" with "Open Help," which opens the session help modal.
- **Stats Exclusions:** Vocabulary exclusions now persist in the immersion database and migrate existing browser-local exclusions on first load.
- **Config Defaults:** Disabled texthooker startup, subtitle, and annotation websocket servers by default. Fresh installs now use a Japanese font stack, transparent subtitle backgrounds, stronger text shadows, and teal N4/fourth-band coloring for primary subtitles. Yomitan popup auto-pause remains enabled.
- **Config Example:** The generated example config now lists every built-in keybinding default.

### Fixed

- **Subtitle Annotations — Grammar Filtering:** Suppressed N+1, JLPT, frequency, and name styling on grammar-only tokens: standalone interjections (`あ`, katakana variants), kana grammar helpers (`ことに`), auxiliary inflection fragments (`れる`, `れた`), polite copula tails (`です`, `じゃないですか`), standalone particles matched by known-word decks, and existence verbs (`ある`/`有る`). Known-word highlighting is preserved where applicable.
- **Subtitle Annotations — Color Priority:** Fixed token color priority so typography settings are preserved, JLPT colors no longer override higher-priority known-word or frequency colors, JLPT underlines persist at their correct color after dictionary lookups and when a token also carries known-word or frequency annotations, and frequency highlighting works correctly for ordinal prefix-noun tokens like `第二`.
- **Subtitle Annotations — Other:** Stopped kana-only tokens from being selected as N+1 targets; preserved Yomitan compound tokens so known component words no longer color a larger unknown word green; kept annotation prefetch running after immediate cache-hit renders; added a brightness lift for annotated token hover states when hover backgrounds are transparent; accepted `subtitleStyle.hoverBackground` as an alias for `subtitleStyle.hoverTokenBackgroundColor`; refreshed the current subtitle after successful card mining so newly known words recolor immediately.
- **Subtitle Bar:** Changed `v` to cycle the primary subtitle bar through visible, hover, and hidden modes with OSD feedback. Added `subtitleStyle.primaryDefaultMode` to set the startup visibility default independently from secondary subtitles.
- **Tokenizer:** Now uses Yomitan `wordClasses` metadata for part-of-speech filtering, and backfills blank MeCab POS fields during parser enrichment.
- **Overlay (Linux):** Fixed multi-line subtitle copy timing out after the prompt; follow-up number-row digits are now accepted for multi-line mining even when the original shortcut modifiers are still held.
- **Overlay (Hyprland):** Fixed fullscreen transitions so overlay geometry refreshes on mpv fullscreen changes, topmost stacking is reasserted, and hover pause works correctly after resize/toggle cycles. Overlay windows now align precisely to mpv bounds with floating decoration disabled; the stats overlay is opaque to prevent mpv bleed-through at the top edge; overlay windows no longer pin across workspaces.
- **Overlay (macOS):** Kept the overlay visible and interactive during transient tracker refreshes while mpv is the active tracked window, and kept it behind unrelated foreground windows while remaining above mpv.
- **Overlay:** Keyboard-only Yomitan popup shortcuts now take precedence over overlay keybindings like `j`; the browser focus outline is hidden so focused overlays no longer show a yellow/orange viewport border.
- **Default Keybindings:** Fixed replay/next subtitle keybindings — session help moved to `Ctrl/Cmd+/`, freeing `Ctrl+Shift+H` and `Ctrl+Shift+L` for subtitle playback controls. `Ctrl+Shift+L` now correctly reaches play-next-subtitle, and play-next resumes from a paused state before pausing again at the subtitle end.
- **Anki:** Manual clipboard subtitle updates preserve existing word audio while replacing sentence audio, animated-image media, and expression fields — even when audio overwrite is configured off.
- **AniList:** Post-watch progress checks now run on time-position updates using the fresh mpv position; manual mark-watched forces a progress sync; missing episode metadata is filled from the filename parser. Duplicate writes during concurrent checks are prevented, and manual watched marks are preserved when sync fails.
- **AniList (Linux):** Retried safeStorage availability after transient keyring failures so tokens can load and save once the keyring becomes available. Prevented config reload from opening the setup window during playback when token storage cannot be resolved, and stopped the setup flow from reporting success when token persistence fails.
- **mpv:** Stopped mpv from holding SubMiner subprocesses during shutdown, preventing desktop crash notifications on video close. Kept the overlay alive across same-media buffering reloads to avoid duplicate startup gates and AniSkip lookups; playlist navigation now reuses the running overlay without repeating the pause-until-ready warmup gate.
- **Launcher:** Managed playback now exits the background SubMiner app when the video closes; explicit background launches remain persistent.
- **Stats:** Background mode routes through the isolated stats daemon; app startup defers to an already-running daemon instead of failing when the port is already in use. Fixed recent session detail pages showing "Media not found" before lifetime media summaries are available.
- **Jellyfin:** Improved setup with recent server selection and inline authentication feedback. Added a tray toggle for runtime-only cast discovery.

### Docs

- Improved the docs homepage with canonical URLs and a cleaner sitemap.

## Installation

See the README and docs/installation guide for full setup steps.

## Assets

- Linux: `SubMiner.AppImage`
- macOS: `SubMiner-*.dmg` and `SubMiner-*.zip`
- Optional extras: `subminer-assets.tar.gz` and the `subminer` launcher

Note: the `subminer` wrapper script uses Bun (`#!/usr/bin/env bun`), so `bun` must be installed and on `PATH`.
