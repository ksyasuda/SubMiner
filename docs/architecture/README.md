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

Packaged apps include a private Bun runtime. First-run setup can install a managed `subminer` wrapper that invokes it by absolute path. macOS wrappers reference app resources. Windows stages a versioned private Bun copy under `%LOCALAPPDATA%\SubMiner\launcher-runtime` so app updates can replace the bundle while a launcher is running; its wrapper still references the bundled launcher script. Linux stages a persistent runtime and launcher copy under `${XDG_DATA_HOME:-~/.local/share}/SubMiner/launcher` so an AppImage launcher survives after the AppImage exits. Desktop startup refreshes managed payloads after an app version change. Standalone launcher assets and development keep using system Bun.

## Read Next

- [Domains](./domains.md) - who owns what
- [Layering](./layering.md) - how modules should depend on each other
- [Subtitle Overlay Priming](./subtitle-overlay-priming.md) - visible-overlay subtitle startup flow
- Public contributor summary: [`docs-site/architecture.md`](../../docs-site/architecture.md)

## Current Shape

- `src/main/` owns composition, runtime setup, IPC wiring, and app lifecycle adapters.
- `src/main/boot/` owns boot-phase assembly seams so `src/main.ts` can stay focused on lifecycle coordination and startup-path selection.
- `src/core/services/` owns focused runtime services plus pure or side-effect-bounded logic.
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
