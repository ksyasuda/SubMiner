<!-- read_when: adding dependencies, moving files, or reviewing architecture drift -->

# Layering Rules

Status: active  
Last verified: 2026-10-31  
Owner: Kyle Yasuda  
Read when: deciding whether a dependency direction is acceptable

## Preferred Dependency Flow

1. `src/main.ts`
2. `src/main/` composition and runtime adapters
3. `src/core/services/` focused services
4. `src/core/utils/` and other pure helpers

Renderer, launcher, plugin, and stats each keep their own local layering and should not become a grab bag for unrelated cross-runtime behavior.

## Rules

- Keep `src/main.ts` thin; wire, do not implement.
- Pass deps straight to a handler factory (`createXHandler({...})`). Do not add
  deps builders (`createBuild*MainDepsHandler`, `*-main-deps.ts`) that re-wrap deps.
  Adaptation logic (defaults, appState getters/setters, shared state) belongs in the
  handler or in a named adapter that returns the deps directly.
- Group a domain's runtimes in one assembly under `src/main/assemblies/` (for example
  `createYoutubeAssembly(deps)`). It owns the domain's state, takes an explicit deps
  object, and returns its runtimes; `main.ts` builds it once and calls through the result.
  Siblings built later in `main.ts` are passed as getters or callbacks, and construction must
  not call them.
- An assembly that answers renderer IPC or CLI commands returns them as typed slices
  (`IpcSlice<...>` / `CliSlice<...>` from `src/main/assemblies/slices.ts`). `main.ts`
  spreads the slices into the registration and checks the result with
  `satisfies Required<...>`, so a missing handler fails typecheck. Only app-level handlers
  with no owning assembly stay inline in `main.ts`.
- Electron modules that value-import named Electron APIs (`dialog`, `screen`, `shell`,
  `clipboard`, `globalShortcut`, `Tray`, ...) cannot load under `bun test`. Pass those APIs
  into an assembly as deps instead of importing them there.
- Prefer injecting dependencies from `src/main/` instead of reaching outward from core services.
- Keep side effects explicit and close to composition boundaries.
- Put reusable business logic in focused services, not in top-level lifecycle files.
- Keep renderer concerns in `src/renderer/`; avoid leaking DOM behavior into main-process code.
- Treat `launcher/*.ts` and the runtime bootstrap generators as source of truth for launcher behavior. `scripts/build-launcher.ts` creates `dist/launcher`; never hand-edit those artifacts.

## Smells

- `main.ts` grows because logic was not extracted
- service reaches directly into unrelated runtime state
- renderer code depends on main-process internals
- docs-site page becomes the only place internal architecture is explained
