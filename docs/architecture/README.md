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
- `src/core/services/` owns focused runtime services plus pure or side-effect-bounded logic.
- `src/core/services/subtitle-generation*.ts` shares local whisper.cpp transcription, safe model downloads, and progress between the launcher and Electron. Optional Silero detection groups short speech passages before transcription, decodes each independently, and restores original media timing without joining omitted gaps. `src/main/runtime/subtitle-generation-runtime.ts` owns the overlay job lifecycle and only loads completed subtitles into the same local media; `src/shared/subtitle-generation*.ts` owns configuration, the multilingual model catalog, and IPC contracts. The overlay runtime retains a session model selection, validates picker requests through IPC, and keeps external model paths authoritative.
- `src/renderer/` owns overlay rendering and input behavior.
- `src/config/` owns config definitions, defaults, loading, and resolution.
- `src/types/` owns shared cross-runtime contracts via domain entrypoints; `src/types.ts` stays a compatibility barrel.
- `src/main/runtime/composers/` owns larger domain compositions.
- `src/main.ts` call sites invoke configured runtime handlers and runtime-object methods directly;
  do not add local pass-through wrappers around them.

## Architecture Intent

- Small units, explicit boundaries
- Composition over monoliths
- Pure helpers where possible
- Stable user behavior while internals evolve

Startup resolves and creates the user-data directory in `src/main-entry-runtime.ts`
before the entry process requests Electron's single-instance lock. Main-process
config bootstrap then writes the default config only when no config file exists.
