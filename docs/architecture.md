# Architecture

SubMiner is migrating from a single, monolithic `src/main.ts` runtime toward a composable architecture with clear extension points.

## Current Structure

- `src/main.ts`: bootstrap/composition root plus remaining legacy orchestration.
- `src/core/`: shared runtime primitives:
  - `app-orchestrator.ts`: lifecycle wiring for ready/activate/quit hooks.
  - `action-bus.ts`: typed action dispatch path.
  - `actions.ts`: canonical app action types.
  - `module.ts` / `module-registry.ts`: module lifecycle contract.
  - `services/`: extracted runtime services (IPC, shortcuts, subtitle, overlay, MPV command routing).
- `src/ipc/`: shared IPC contract + wrappers used by main, preload, and renderer.
- `src/subtitle/`: staged subtitle pipeline (`normalize` -> `tokenize` -> `merge`).
- `src/modules/`: feature modules:
  - `runtime-options`
  - `jimaku`
  - `subsync`
  - `anki`
  - `texthooker`
- provider registries:
  - `src/window-trackers/index.ts`
  - `src/tokenizers/index.ts`
  - `src/token-mergers/index.ts`
  - `src/translators/index.ts`
  - `src/subsync/engines.ts`

## Migration Status

- Completed:
  - Action bus wired for CLI, shortcuts, and IPC-triggered commands.
  - Command/action mapping covered by focused core tests.
  - Shared IPC channel contract adopted across main/preload/renderer.
  - Runtime options extracted into module lifecycle.
  - Provider registries replace hardcoded backend selection.
  - Subtitle tokenization/merge/enrich flow moved to staged pipeline.
  - Stage-level subtitle pipeline tests added for deterministic behavior.
  - Jimaku, Subsync, Anki, and texthooker/websocket flows moduleized.
- In progress:
  - Further shrink `src/main.ts` by moving orchestration into dedicated services/orchestrator files.
  - Continue moduleizing remaining integrations with complex lifecycle coupling.

## Design Rules

- New feature behavior should be added as:
  - a module (`src/modules/*`) and/or
  - a provider registration (`src/*/index.ts` registries),
  instead of editing unrelated branches in `src/main.ts`.
- New command triggers should dispatch `AppAction` and reuse existing action handlers.
- New IPC channels should be added only via `src/ipc/contract.ts`.
