<!-- read_when: changing runtime wiring, moving code across layers, or trying to find ownership -->

# Architecture Map

Status: active
Last verified: 2026-03-31
Owner: Kyle Yasuda
Read when: runtime ownership, composition boundaries, or layering questions

SubMiner runs as three cooperating runtimes:

- Electron desktop app in `src/`
- Launcher CLI in `launcher/`
- mpv Lua plugin in `plugin/subminer/`

The desktop app keeps `src/main.ts` as composition root and pushes behavior into small runtime/domain modules.

## Read Next

- [Domains](./domains.md) - who owns what
- [Layering](./layering.md) - how modules should depend on each other
- Public contributor summary: [`docs-site/architecture.md`](../../docs-site/architecture.md)

## Current Shape

- `src/main/` owns composition, runtime setup, IPC wiring, and app lifecycle adapters.
- `src/main/*.ts` wrapper runtimes sit between `src/main.ts` and `src/main/runtime/**`
  so the composition root stays thin while exported `createBuild*MainDepsHandler`
  APIs remain internal plumbing. Key domain runtimes:
  - `anilist-runtime` – AniList token management, media tracking, retry queue
  - `cli-startup-runtime` – CLI command dispatch and initial-args handling
  - `discord-presence-lifecycle-runtime` – Discord Rich Presence lifecycle
  - `first-run-runtime` – first-run setup wizard
  - `ipc-runtime` – IPC handler registration and composition
  - `jellyfin-runtime` – Jellyfin session, playback, mpv orchestration
  - `main-startup-runtime` – top-level startup orchestration (app-ready → CLI → headless)
  - `main-startup-bootstrap` – wiring helper that builds startup runtime inputs
  - `mining-runtime` – Anki card mining actions
  - `mpv-runtime` – mpv client lifecycle
  - `overlay-ui-runtime` – overlay window management, visibility, tray
  - `overlay-geometry-runtime` – overlay bounds resolution
  - `shortcuts-runtime` – global shortcut registration
  - `startup-sequence-runtime` – headless known-word refresh and deferred startup sequencing
  - `stats-runtime` – immersion tracker, stats server, stats CLI
  - `subtitle-runtime` – subtitle prefetch, tokenization, caching
  - `youtube-runtime` – YouTube playback flow
  - `yomitan-runtime` – Yomitan extension loading and settings
- `src/main/boot/` owns boot-phase assembly seams so `src/main.ts` can stay focused on lifecycle coordination and startup-path selection.
- `src/core/services/` owns focused runtime services plus pure or side-effect-bounded logic.
- `src/renderer/` owns overlay rendering and input behavior.
- `src/config/` owns config definitions, defaults, loading, and resolution.
- `src/types/` owns shared cross-runtime contracts via domain entrypoints; `src/types.ts` stays a compatibility barrel.
- `src/main/runtime/composers/` owns larger domain compositions.

## Architecture Intent

- Small units, explicit boundaries
- Composition over monoliths
- Pure helpers where possible
- Stable user behavior while internals evolve
