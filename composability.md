# SubMiner Composability and Extensibility Plan

## Goals

- Reduce coupling concentrated in `src/main.ts`.
- Make new features additive (new files/modules) instead of invasive (edits across many files).
- Standardize extension points for trackers, subtitle processing, IPC actions, and integrations.
- Preserve existing behavior while incrementally migrating architecture.

## Current Constraints (From Existing Code)

- `src/main.ts` is the effective integration bus (~5k LOC) and mixes lifecycle, IPC, shortcuts, mpv integration, overlay control, and feature flows.
- Input surfaces are fragmented (CLI args, global shortcuts, renderer IPC) and often route through separate logic paths.
- Some extension points already exist (`src/window-trackers/*`, centralized config definitions in `src/config/definitions.ts`) but still use hardcoded selection/wiring.
- Type contracts are duplicated between main/preload/renderer in places, which increases drift risk.

## Target Architecture

### 1. Composition Root + Services

- Keep `src/main.ts` as bootstrap only.
- Move behavior into focused services:
  - `src/core/app-orchestrator.ts`
  - `src/core/services/overlay-service.ts`
  - `src/core/services/mpv-service.ts`
  - `src/core/services/shortcut-service.ts`
  - `src/core/services/ipc-service.ts`
  - `src/core/services/subtitle-service.ts`

### 2. Module System

- Introduce module contract:
  - `src/core/module.ts`
  - `src/core/module-registry.ts`
- Built-in features become modules:
  - anki module
  - jimaku module
  - runtime options module
  - texthooker/websocket module
  - subsync module

### 3. Action Bus

- Add typed action dispatch (single command path for CLI/shortcut/IPC):
  - `src/core/actions.ts` (action type map)
  - `src/core/action-bus.ts` (register/dispatch)
- All triggers emit actions; business logic subscribes once.

### 4. Provider Registries (Strategy Pattern)

- Replace hardcoded switch/if wiring with registries:
  - tracker providers
  - tokenizer providers
  - translation providers
  - subsync backend providers

### 5. Shared IPC Contract

- Single source of truth for IPC channels and payloads:
  - `src/ipc/contract.ts`
  - `src/ipc/main-api.ts`
  - `src/ipc/renderer-api.ts`
- `preload.ts` and renderer consume typed wrappers instead of ad hoc channel strings.

### 6. Subtitle Pipeline

- Middleware/stage pipeline:
  - ingest -> normalize -> tokenize -> merge -> enrich -> emit
- Files:
  - `src/subtitle/pipeline.ts`
  - `src/subtitle/stages/*.ts`

### 7. Module-Owned Config Extensions

- Keep `src/config/definitions.ts` as the central resolved output.
- Add registration path so modules can contribute config/runtime option metadata and defaults before final resolution.

## Phased Delivery Plan

## Phase 0: Baseline and Safety Net

### Scope

- Add architecture docs and migration guardrails.
- Increase tests around currently stable behavior before structural changes.

### Changes

- Add architecture decision notes:
  - `docs/development.md` (new section: architecture/migration)
  - `docs/architecture.md` (new)
- Add basic smoke tests for command routing, overlay visibility toggles, and config load behavior.

### Exit Criteria

- Existing behavior documented with acceptance checklist.
- CI/build + config tests green.

## Phase 1: Extract I/O Surfaces

### Scope

- Isolate IPC and global shortcut registration from `src/main.ts` without changing semantics.

### Changes

- Create:
  - `src/core/services/ipc-service.ts`
  - `src/core/services/shortcut-service.ts`
- Move `ipcMain.handle/on` registration blocks from `src/main.ts` into `IpcService.register(...)`.
- Move global shortcut registration into `ShortcutService.registerGlobal(...)`.
- Keep old entrypoints in `main.ts` delegating to new service methods.

### Exit Criteria

- No user-visible behavior changes.
- `src/main.ts` shrinks materially.
- Manual verification:
  - overlay toggles
  - copy/mine shortcuts
  - runtime options open/cycle

## Phase 2: Introduce Action Bus

### Scope

- Canonicalize command execution path.

### Changes

- Add typed action bus (`src/core/action-bus.ts`, `src/core/actions.ts`).
- Convert these triggers to dispatch actions instead of direct calls:
  - CLI handling (`handleCliCommand` path)
  - shortcut handlers
  - IPC command handlers
- Keep existing business functions as subscribers initially.

### Exit Criteria

- One action path per command (no duplicated behavior branches).
- Unit tests for command mapping (CLI -> action, shortcut -> action, IPC -> action).

## Phase 3: Module Runtime Skeleton

### Scope

- Add module contract and migrate one low-risk feature first.

### Changes

- Introduce:
  - `src/core/module.ts`
  - `src/core/module-registry.ts`
  - `src/core/app-context.ts`
- First migration target: runtime options
  - create `src/modules/runtime-options/module.ts`
  - wire existing `RuntimeOptionsManager` through module lifecycle.

### Exit Criteria

- Runtime options fully owned by module.
- Core app can start with module list and deterministic module init order.

## Phase 4: Provider Registries

### Scope

- Remove hardcoded provider selection.

### Changes

- Window tracker:
  - Replace switch in `src/window-trackers/index.ts` with registry API.
  - Add provider objects for hyprland/sway/x11/macos.
- Tokenizer/translator/subsync:
  - Add analogous registries in new provider directories.

### Exit Criteria

- Adding a provider requires adding one file + registration line.
- No edits required in composition root for new provider variants.

## Phase 5: Shared IPC Contract

### Scope

- Eliminate channel/payload drift between main/preload/renderer.

### Changes

- Add typed IPC contract files under `src/ipc/`.
- Refactor `src/preload.ts` to generate API from shared contract wrappers.
- Refactor renderer calls to consume typed API interface from shared module.

### Exit Criteria

- Channel names declared once.
- Compile-time checking across main/preload/renderer for payload mismatch.

## Phase 6: Subtitle Pipeline

### Scope

- Extract subtitle transformations into composable stages.

### Changes

- Create pipeline framework and migrate existing tokenization/merge flow:
  - stage: normalize subtitle
  - stage: tokenize (`MecabTokenizer` adapter)
  - stage: merge tokens (`mergeTokens` adapter)
  - stage: post-processing/enrichment hooks
- Integrate pipeline execution into subtitle update loop.

### Exit Criteria

- Current output parity for tokenization/merged token behavior.
- Stage-level tests for deterministic input/output.

## Phase 7: Moduleized Integrations

### Scope

- Convert major features to modules in descending complexity.

### Migration Order

1. Jimaku
2. Texthooker/WebSocket bridge
3. Subsync
4. Anki integration

### Exit Criteria

- Each feature has independent init/start/stop lifecycle.
- App startup composes modules instead of hardcoded inline setup.

## Recommended File Layout (Target)

```text
src/
  core/
    app-orchestrator.ts
    app-context.ts
    actions.ts
    action-bus.ts
    module.ts
    module-registry.ts
    services/
      ipc-service.ts
      shortcut-service.ts
      overlay-service.ts
      mpv-service.ts
      subtitle-service.ts
  modules/
    runtime-options/module.ts
    jimaku/module.ts
    texthooker/module.ts
    subsync/module.ts
    anki/module.ts
  ipc/
    contract.ts
    main-api.ts
    renderer-api.ts
  subtitle/
    pipeline.ts
    stages/
      normalize.ts
      tokenize-mecab.ts
      merge-default.ts
      enrich.ts
```

## PR Breakdown (Suggested)

1. PR1: Phase 1 (`IpcService`, `ShortcutService`, no behavior changes)
2. PR2: Phase 2 (action bus + trigger mapping)
3. PR3: Phase 3 (module skeleton + runtime options module)
4. PR4: Phase 4 (window tracker registry + provider pattern)
5. PR5: Phase 5 (shared IPC contract)
6. PR6: Phase 6 (subtitle pipeline)
7. PR7+: Phase 7 (feature-by-feature module migrations)

## Validation Matrix

- Build/typecheck:
  - `pnpm run build`
- Config correctness:
  - `pnpm run test:config`
- Manual checks per PR:
  - app start/stop/toggle CLI
  - visible/invisible overlay control
  - global shortcuts
  - subtitle render and token display
  - runtime options open + cycle
  - anki mine flow + update-from-clipboard

## Backward Compatibility Rules

- Preserve existing CLI flags and IPC channel behavior during migration.
- Maintain config shape compatibility; only additive changes unless versioned migration is added.
- Keep default behavior equivalent for all supported compositor backends.

## Key Risks and Mitigations

- Risk: behavior regressions during extraction.
  - Mitigation: move code verbatim first, refactor second; keep thin delegates in `main.ts` until stable.
- Risk: module lifecycle race conditions.
  - Mitigation: explicit init order, idempotent `start/stop`, and startup dependency checks.
- Risk: IPC contract churn breaks renderer.
  - Mitigation: contract wrappers and compile-time types; temporary compatibility adapters.
- Risk: feature coupling around anki/mine flows.
  - Mitigation: defer high-coupling module migrations until action bus and shared app context are stable.

## Definition of Done (Program-Level)

- `src/main.ts` reduced to bootstrap/composition responsibilities.
- New feature prototype can be added as an isolated module with:
  - module file
  - optional provider registration
  - optional config schema registration
  - no invasive edits across unrelated subsystems
- IPC and command routing are single-path and typed.
- Existing user workflows remain intact.
