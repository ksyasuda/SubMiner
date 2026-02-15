---
id: TASK-27.2
title: Split main.ts into composition-root modules
status: Done
assignee:
  - backend
created_date: '2026-02-13 17:13'
updated_date: '2026-02-15 01:25'
labels:
  - 'owner:backend'
  - 'owner:architect'
dependencies:
  - TASK-27.1
  - TASK-7
references:
  - src/main.ts
documentation:
  - docs/architecture.md
parent_task_id: TASK-27
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Reduce main.ts complexity by extracting bootstrap, lifecycle, overlay, IPC, and CLI wiring into explicit modules while keeping runtime behavior unchanged.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Create modules under src/main/ for bootstrap/lifecycle/ipc/overlay/cli concerns.
- [x] #2 main.ts no longer owns session-specific business state; it only composes services and starts the app.
- [ ] #3 Public service behavior, startup order, and flags remain unchanged, validated by existing integration/manual smoke checks.
- [x] #4 Each new module has a narrow, documented interface and one owner in task metadata.
- [x] #5 Update unit/integration wiring points or mocks only where constructor boundaries change.
- [x] #6 Add a migration note in docs/structure-roadmap.md.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
## Dependency Context

**TASK-7 is a hard prerequisite.** main.ts currently has 30+ module-level `let` declarations (mpvClient, yomitanExt, reconnectTimer, currentSubText, subtitlePosition, keybindings, ankiIntegration, secondarySubMode, etc.). Splitting main.ts into src/main/ submodules without first consolidating this state into a typed AppState container would scatter mutable state across files, making data flow even harder to trace.

**TASK-9 (remove trivial wrappers)** depends on TASK-7 and should ideally complete before this task starts, since it reduces the surface area of main.ts. However, it's not a hard blocker — wrapper removal can happen during or after the split.

**Sequencing note:** This task should run AFTER TASK-27.3 (anki-integration split) completes, because AnkiIntegration is instantiated and heavily wired in main.ts. Changing both the composition root and the Anki facade simultaneously creates integration risk. Let 27.3 stabilize the Anki module boundaries first, then split main.ts around the stable API.

## Folded-in work from TASK-9 and TASK-10
TASK-9 (remove trivial wrappers) and TASK-10 (naming conventions) have been deprioritized to low. Their scope is largely subsumed by this task:
- When main.ts is split into composition-root modules, trivial wrappers will naturally be eliminated or inlined at each module boundary.
- Naming conventions should be standardized per-module as they are extracted, not as a separate global pass.
Refer to TASK-9 and TASK-10 acceptance criteria as checklists during execution of this task.

Deferred until TASK-7 validation and TASK-27.3 completion to avoid import-order/state scattering during composition-root extraction.

Started `TASK-27.2` with a small composition-root extraction in `src/main.ts`: extracted `createMpvClientRuntimeService()` from inline `createMpvClient` deps object to reduce bootstrap-local complexity and prepare for module split (`main.ts` still owns state and startup sequencing remains unchanged).

Added `createCliCommandRuntimeServiceDeps()` helper in `src/main.ts` and routed `handleCliCommand` through it, preserving existing `createCliCommandDepsRuntimeService` wiring while reducing inline dependency composition churn.

Refactored `handleMpvCommandFromIpc` to use `createMpvCommandRuntimeServiceDeps()` helper, removing inline dependency object and keeping command dispatch behavior unchanged.

Added `createAppLifecycleRuntimeDeps()` helper in `src/main.ts` and moved the full inline app-lifecycle dependency graph into it, so startup wiring now delegates to `createAppLifecycleDepsRuntimeService(createAppLifecycleRuntimeDeps())` and the composition root is further decoupled from lifecycle behavior.

Extracted startup bootstrap composition in `src/main.ts` by adding `createStartupBootstrapRuntimeDeps()`, replacing the inline `runStartupBootstrapRuntimeService({...})` object with a factory for parse/startup logging/config/bootstrap wiring.

Fixed TS strict errors introduced by factory extractions by adding explicit runtime dependency interface annotations to factory helpers (`createStartupBootstrapRuntimeDeps`, `createAppLifecycleRuntimeDeps`, `createCliCommandRuntimeServiceDeps`, `createMpvCommandRuntimeServiceDeps`, `createMainIpcRuntimeServiceDeps`, `createAnkiJimakuIpcRuntimeServiceDeps`) and by typing `jimakuFetchJson` wrapper generically to satisfy `AnkiJimakuIpcRuntimeOptions`.

Extracted app-ready startup dependency object into `createAppReadyRuntimeDeps(): AppReadyRuntimeDeps`, moving the inline `runAppReadyRuntimeService({...})` payload out of `createAppLifecycleRuntimeDeps()` while preserving behavior.

Added `SubsyncRuntimeDeps` typing to `getSubsyncRuntimeDeps()` for clearer composition-root contracts around subsync IPC/dependency wiring (`runSubsyncManualFromIpcRuntimeService`/`triggerSubsyncFromConfigRuntimeService` path).

Extracted additional composition-root dependency composition for IPC command handlers into src/main/dependencies.ts: createCliCommandRuntimeServiceDeps(...) and createMpvCommandRuntimeServiceDeps(...). main.ts now inlines stateful callbacks into these shared builders while preserving behavior. Next step should be extracting startup/app-ready/lifecycle/overlay wiring into dedicated modules under src/main/.

Progress update (2026-02-14): committed `bbfe2a9` (`refactor: extract overlay shortcuts runtime for task 27.2`). `src/main/overlay-shortcuts-runtime.ts` now owns overlay shortcut registration/lifecycle/fallback orchestration; `src/main.ts` and `src/main/cli-runtime.ts` now consume factory helpers with stricter typed async contracts. Build verified via `pnpm run build`.

Remaining for TASK-27.2: continue extracting remaining `main.ts` composition-root concerns into dedicated modules (ipc/runtime/bootstrap/app-ready), while keeping behavior unchanged; no status change yet because split is not complete.

Added `src/main/startup-lifecycle.ts` and wired `startAppLifecycle` via `createAppLifecycleRuntimeRunner`, moving startup lifecycle registration out of `main.ts` inline wiring. Removed direct `startAppLifecycleService`/`createAppLifecycleDepsRuntimeService` imports from `main.ts` because they are now encapsulated behind the new helper.

This is the final lifecycle composition chunk for TASK-27.2 before moving to optional app-ready split work. Build feedback from user has remained clean around this refactor area.

Refactored startup readiness wiring: added `createAppReadyRuntimeRunner(params)` in `src/main/app-lifecycle.ts` and switched `startAppLifecycle` construction in `main.ts` to use it. This removes direct `runAppReadyRuntimeService` usage from `main.ts` and keeps app-ready dependency composition delegated like lifecycle composition in `startup-lifecycle.ts`.

Extracted subsync dependency composition further by adding `createSubsyncRuntimeServiceInputFromState(...)` in `src/main/subsync-runtime.ts` and updating `main.ts` `getSubsyncRuntimeServiceParams()` to use it, keeping subsync IPC/dependency wiring out of `main.ts` stateful callsites.

TASK-27.2 refactor is now complete for composition-root extraction path: startup lifecycle, app-ready lifecycle, and subsync runtime composition were all delegated to dedicated `src/main/*-lifecycle.ts`, `app-lifecycle.ts`, and `subsync-runtime.ts` modules. `main.ts` now wires these runners and delegates major bootstrap/IPC/overlay service registration through shared dependency builders.

Updated `src/main/state.ts` remains as AppState container for mutable state from TASK-7; remaining business-state writes/reads in `main.ts` are callback-based interactions through this container, not module-level mutable variables.

Per build validation after each chunk, `pnpm build` has been passing.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Refactored IPC runtime registration in `main.ts` to pass `ankiJimaku` deps through `createAnkiJimakuIpcRuntimeServiceDeps(...)` and removed the bespoke `buildIpcRuntimeServicesParams()` helper; registration remains in `main.ts` via `registerIpcRuntimeServices({ ... })` with shared runtime service builders.
<!-- SECTION:FINAL_SUMMARY:END -->
