<!-- read_when: changing runtime wiring, moving code across layers, or trying to find ownership -->

# Architecture Map

Status: active
Last verified: 2026-05-23
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
