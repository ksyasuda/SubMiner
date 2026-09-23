<!-- read_when: changing runtime wiring, moving code across layers, or trying to find ownership -->

# Architecture Map

Status: active
Last verified: 2026-05-23
Owner: Kyle Yasuda
Read when: runtime ownership, composition boundaries, or layering questions

SubMiner runs as three cooperating runtimes:

- Electron desktop app in `src/`
- Launcher CLI in `launcher/`, with managed app-installed wrappers in `src/main/runtime/managed-launcher.ts`
- mpv Lua plugin in `plugin/subminer/`

The desktop app keeps `src/main.ts` as composition root and pushes behavior into small runtime/domain modules.

Packaged apps include a private Bun runtime. Setup and release assets provide bootstrap wrappers generated from `src/main/runtime/*-launcher-bootstrap.ts`. macOS runs Bun and the CLI from app resources. Windows stages a versioned private Bun copy under `%LOCALAPPDATA%\SubMiner\launcher-runtime/<version>` so a running launcher does not lock the updater-owned app executable. Linux stages Bun, the matching CLI, and licenses under `${XDG_DATA_HOME:-~/.local/share}/SubMiner/launcher`. Its steady-state path performs one app `stat` and starts the cache without Electron. A missing cache or changed app fingerprint runs `launcher/prepare.cjs` through Electron's Node mode to refresh it. Desktop startup migrates recognized writable legacy JavaScript launchers and refreshes managed payloads after app changes. Development commands still use system Bun.

Update checks and startup launcher migration share a serialized update-state store. Deferred launcher paths are acknowledged only after migration succeeds or the candidate is no longer eligible. Running Windows launchers and unreadable or unwritable candidates remain pending for a later startup.

## Read Next

- [Domains](./domains.md) - who owns what
- [Layering](./layering.md) - how modules should depend on each other
- [Subtitle Overlay Priming](./subtitle-overlay-priming.md) - visible-overlay subtitle startup flow
- Public contributor summary: [`docs-site/architecture.md`](../../docs-site/architecture.md)

## Current Shape

- `src/main/` owns composition, runtime setup, IPC wiring, and app lifecycle adapters.
- `src/main/boot/` owns boot-phase assembly seams so `src/main.ts` can stay focused on lifecycle coordination and startup-path selection.
- `src/main/runtime/linux-overlay-mode-runtime.ts` owns Linux fullscreen mode state and window replacement. App cleanup cancels pending replacements; `main.ts` supplies window creation and subtitle refresh hooks.
- `src/core/services/` owns focused runtime services plus pure or side-effect-bounded logic.
- `src/core/services/subtitle-generation*.ts` shares local whisper.cpp transcription, safe model downloads, and progress between the launcher and Electron. Optional dialogue mode retains both Silero-detected speech and other audible sections, omits confidently silent gaps, decodes passages independently, and restores original media timing. `src/main/runtime/subtitle-generation-runtime.ts` owns the overlay job lifecycle and only loads completed subtitles into the same local media; `src/shared/subtitle-generation*.ts` owns configuration, the multilingual model catalog, and IPC contracts. The overlay runtime retains a session model selection, validates picker requests through IPC, and keeps external model paths authoritative.
- Subtitle model recommendations use bounded `nvidia-smi` and Whisper CUDA discovery probes in `subtitle-generation-acceleration.ts`. The overlay runtime caches results by executable path for 30 seconds and exposes acceleration status through the existing status IPC. Recommendations do not alter model selection or transcription arguments.
- `subtitle-generation-reference.ts` ranks mpv's loaded text subtitle tracks, excludes signs/songs and forced references, and extracts timing hints with FFmpeg. The overlay and launcher snapshot references only for matching media, including active subtitle delays. Hints guide long-passage cuts with or without VAD; they never limit audio coverage or replace Whisper timestamps.
- `src/main/runtime/subtitle-selection.ts` reads and validates mpv subtitle tracks and applies primary/secondary selections. Its opt-in session shortcut opens the shared overlay modal window, with renderer focus and subtitle suppression handled by the modal registry.
- `src/shared/session-key-sequences.ts` rejects sequence prefixes reserved by single-key actions. The session-binding compiler reserves configured and built-in overlay keys; `src/main/runtime/session-bindings-runtime.ts` adds active mpv bindings and publishes the effective list to both the plugin artifact and the renderer through `session-bindings:changed`. mpv no-op `ignore` bindings do not reserve prefixes.
- `src/renderer/` owns overlay rendering and input behavior.
- `src/config/` owns config definitions, defaults, loading, and resolution.
- `src/types/` owns shared cross-runtime contracts via domain entrypoints; `src/types.ts` stays a compatibility barrel.
- `src/main/runtime/composers/` owns larger domain compositions.
- `src/main.ts` call sites invoke configured runtime handlers and runtime-object methods directly;
  do not add local pass-through wrappers around them.

## Architecture Intent

The dictionary backend is selected once at startup by `dictionaryBackend`. Yomitan keeps its existing session and external-profile policy. Hachidori uses `persist:hachidori`; overlay windows select that session before extension loading, including deferred startup. Named settings flags can open either backend without injecting a second reader into the active overlay. The detached stats word helper reads the same config key so dashboard mining uses the active backend.

`setup-state.json` records one backend's status at a time plus `completedDictionaryBackends`, the backends that finished setup before. The app projects the file onto its active backend on startup and stamps that backend into the file. The launcher gates playback on the stamped backend when an app is already running, since a config edit takes effect only after restart.

`vendor/hachidori/` is a submodule of `ksyasuda/hachidori`, tracking the `subminer` branch and pinned to a tested commit. Its nested HoshiDicts submodule and WASM binaries remain upstream versions. Initialize sources with `git submodule update --init --recursive`; merge upstream updates in the fork, test them, then update SubMiner's submodule commit. `SOURCE.json` records the upstream base and artifact checksums; the submodule commit identifies the integrated version. `build:hachidori` verifies recorded artifact checksums and stages the extension for development and packaging. It enables overlay mode, disables custom JavaScript, keeps the lookup highlight on in the overlay first-install options (SubMiner captures media from mpv, not the overlay viewport), removes the unsupported `userScripts` permission, and sets a fixed manifest `key` so the extension ID (and the storage origin holding its dictionaries and settings) does not depend on the userData path, only in that staged copy; the fork keeps upstream browser defaults. Before loading the extension, its session clears service worker registrations so Electron uses the current bundled code; dictionary databases and settings remain intact. First-run setup uses Hachidori sharing messages to link or unlink external dictionary hosts and checks their live inventory. Linked dictionaries and dictionary edits use the host, while Anki configuration, pronunciation sources, custom buttons, and mining stay local to SubMiner. The parser bridge adapts its runtime messages to the existing subtitle scanner and dictionary automation. Scanning retains term-entry frequencies, and only tokens without ranks need further frequency lookups through the existing term-entry API. This requires a matching definition entry and does not preserve the frequency source's reading provenance. SubMiner consumes native `hachidori-popup-shown` and `hachidori-popup-hidden` attention events for mouse handling, keyboard focus, and the subtitle sidebar. Attention also covers a left press anywhere on the overlay that may start a selection, and the host element only exists after the first lookup, so once a Hachidori event has been seen popup auto-pause requires an unhidden popup pane in the host's shadow root and rechecks after each successful lookup. The fork retains host attributes, hover and successful-lookup notifications, and commands that need private reader state. The Anki proxy strips local duplicate/overwrite metadata before forwarding requests and enriches only confirmed writes.

- Small units, explicit boundaries
- Composition over monoliths
- Pure helpers where possible
- Stable user behavior while internals evolve

Startup resolves and creates the user-data directory in `src/main-entry-runtime.ts`
before the entry process requests Electron's single-instance lock. Main-process
config bootstrap then writes the default config only when no config file exists.
