# Refactoring Investigation Report

## Overview

This report evaluates the SubMiner refactoring effort on the `refactor` branch against the plan defined in `plan.md` and the safety checklist in `docs/refactor-main-checklist.md`. The refactoring aimed to eliminate unnecessary abstraction layers, consolidate related services, fix known bugs, and add test coverage.

**Result**: 65 commits, 104 files changed, main.ts reduced from ~5,800 to 1,398 lines, all 67 tests passing, build succeeds.

---

## Phase 1: Delete Thin Wrappers — COMPLETE

All 9 target wrapper services and 7 associated test files have been deleted and their logic inlined into callers (mostly main.ts or the services they fed).

| Target File | Status |
|-------------|--------|
| `config-warning-runtime-service.ts` + test | Deleted, inlined |
| `overlay-modal-restore-service.ts` + test | Deleted, inlined |
| `runtime-options-manager-runtime-service.ts` + test | Deleted, inlined |
| `app-logging-runtime-service.ts` + test | Deleted, inlined |
| `overlay-send-service.ts` (no test) | Deleted, inlined |
| `startup-resource-runtime-service.ts` + test | Deleted, inlined |
| `config-generation-runtime-service.ts` + test | Deleted, inlined |
| `app-shutdown-runtime-service.ts` + test | Deleted, inlined |
| `shortcut-ui-deps-runtime-service.ts` + test | Deleted, inlined |

**Files removed**: 16 (9 services + 7 tests)
**No issues found.**

---

## Phase 2: Consolidate DI Adapter Services — COMPLETE

All 4 dependency-injection adapter services have been merged into the services they fed, and their test files removed.

| Adapter | Merged Into | Status |
|---------|------------|--------|
| `cli-command-deps-runtime-service.ts` + test | `cli-command-service.ts` | Done |
| `ipc-deps-runtime-service.ts` + test | `ipc-service.ts` | Done |
| `tokenizer-deps-runtime-service.ts` + test | `tokenizer-service.ts` | Done |
| `app-lifecycle-deps-runtime-service.ts` + test | `app-lifecycle-service.ts` | Done |

**Files removed**: 8 (4 services + 4 tests)
**No issues found.**

---

## Phase 3: Consolidate Related Services — COMPLETE

All planned service merges have been executed.

| Consolidation | Source Files | Result File | Status |
|--------------|-------------|-------------|--------|
| Overlay visibility | `overlay-visibility-runtime-service.ts` | `overlay-visibility-service.ts` | Done |
| Overlay broadcast | `overlay-broadcast-runtime-service.ts` + test | `overlay-manager-service.ts` | Done |
| Overlay shortcuts | `overlay-shortcut-lifecycle-service.ts`, `overlay-shortcut-fallback-runner.ts` | `overlay-shortcut-service.ts` + `overlay-shortcut-handler.ts` | Done |
| Numeric shortcuts | `numeric-shortcut-runtime-service.ts` + test, `numeric-shortcut-session-service.ts` | `numeric-shortcut-service.ts` | Done |
| Startup/bootstrap | `startup-bootstrap-runtime-service.ts` + test, `app-ready-runtime-service.ts` + test | `startup-service.ts` | Done |

**Files removed**: ~10
**No issues found.**

---

## Phase 4: Fix Bugs and Code Quality — COMPLETE

### 4.1 Debug console.log statements
**Status**: RESOLVED — No debug `console.log` or `console.warn` calls remain in `overlay-visibility-service.ts`.

### 4.2 `as never` type cast
**Status**: RESOLVED — No `as never` cast remains in `tokenizer-service.ts`. The type mismatch was fixed properly.

### 4.3 fieldGroupingResolver race condition
**Status**: RESOLVED — Fixed in `main.ts` (lines 305–332) with a sequence counter mechanism. Each new resolver increments `fieldGroupingResolverSequence`, and wrapped resolvers check if their sequence matches the current value before executing. Stale resolutions are correctly ignored.

### 4.4 Async callback safety
**Status**: RESOLVED
- **CLI commands** (`cli-command-service.ts`): Async commands use `runAsyncWithOsd` helper (lines 177–187) which catches errors, logs them, and displays via MPV OSD.
- **IPC handlers** (`ipc-service.ts`): Async handlers use `ipcMain.handle` (not `.on`), which properly awaits and propagates promise results. Synchronous handlers correctly use `ipcMain.on`.

### 4.5 `-runtime-service` naming convention
**Status**: RESOLVED — No files with `-runtime-service` in the name exist under `src/core/services/`. All have been renamed to `*-service.ts`.

---

## Phase 5: Add Tests for Critical Untested Services — COMPLETE

Tests added per plan:

| Service | Test File | Tests |
|---------|-----------|-------|
| `mpv-service.ts` (761 lines) | `mpv-service.test.ts` | Socket protocol, property changes, reconnection |
| `subsync-service.ts` (427 lines) | `subsync-service.test.ts` | Config resolution, command construction, error handling |
| `tokenizer-service.ts` (305 lines) | `tokenizer-service.test.ts` | Parser init, token extraction, edge cases |
| `cli-command-service.ts` (204 lines) | `cli-command-service.test.ts` (290 lines) | Expanded: all dispatch paths, error propagation |

**Total**: 67 tests passing across all test suites.

---

## Phase 6: Directory Restructure — NOT STARTED (Optional)

Services remain in the flat `src/core/services/` directory. This phase was explicitly marked optional and low-priority. The current flat structure with ~25 files is manageable.

---

## Checklist Compliance (docs/refactor-main-checklist.md)

### Invariants

| Invariant | Status |
|-----------|--------|
| CLI flags and aliases working | Verified via tests |
| IPC channel names backward-compatible | No channel names changed |
| Overlay toggle behavior preserved | Logic moved verbatim |
| MPV integration behavior preserved | Logic moved verbatim, tested |
| Texthooker mode preserved | Not altered |
| Mining/runtime options paths preserved | Logic moved verbatim |

### Automated Checks

| Check | Status |
|-------|--------|
| `pnpm run build` | Passes |
| `pnpm run test:core` | 67/67 passing |

### Manual Smoke Checks

**Status**: NOT YET PERFORMED — Visual overlay rendering, card mining flow, and field-grouping interaction require a real desktop session with MPV. Automated tests validate logic correctness but cannot catch rendering regressions.

---

## Loose Ends and Concerns

### 1. Unused Architectural Scaffolding (~500 lines, dead code)

The following files exist in the repository but are **not imported or used anywhere**:

**Core abstractions:**
| File | Lines | Purpose |
|------|-------|---------|
| `src/core/action-bus.ts` | 22 | Generic action dispatcher (`register`/`dispatch`) |
| `src/core/actions.ts` | 17 | Union type `AppAction` with 16 action variants |
| `src/core/app-context.ts` | 46 | Module context interfaces |
| `src/core/module-registry.ts` | 37 | Module lifecycle manager (init/start/stop) |
| `src/core/module.ts` | 7 | `SubminerModule<TContext>` interface |

**Module implementations:**
| File | Lines | Purpose |
|------|-------|---------|
| `src/modules/runtime-options/module.ts` | 62 | Runtime options module |
| `src/modules/subsync/module.ts` | 79 | Subsync module |
| `src/modules/jimaku/module.ts` | 73 | Jimaku module |

**IPC abstraction layer:**
| File | Lines | Purpose |
|------|-------|---------|
| `src/ipc/contract.ts` | 62 | IPC channel name constants |
| `src/ipc/main-api.ts` | 20 | Main process IPC helpers |
| `src/ipc/renderer-api.ts` | 28 | Renderer process IPC helpers |

These represent scaffolding for a module-based architecture that was prototyped but never wired into main.ts. The current codebase uses the service-oriented approach throughout.

**Recommendation**: Remove if abandoned, or document with clear intent if planned for a future phase.

### 2. New Consolidated Services Without Test Coverage

Seven service files created or consolidated during Phase 3 lack dedicated tests:

| File | Lines | Risk Level | Reason |
|------|-------|------------|--------|
| `overlay-shortcut-handler.ts` | 216 | Higher | Complex runtime handlers with fallback logic |
| `mining-service.ts` | 179 | Higher | 6 public functions orchestrating mining workflows |
| `anki-jimaku-service.ts` | 173 | Higher | Complex IPC registration with multiple handlers |
| `startup-service.ts` | ~150 | Medium | Bootstrap and app-ready orchestration |
| `numeric-shortcut-service.ts` | 133 | Medium | Session state management |
| `subsync-runner-service.ts` | 86 | Lower | Thin runtime wrapper |
| `jimaku-service.ts` | 81 | Lower | Config accessor functions |

Phase 5 targeted the 4 highest-risk *previously existing* untested services. The Phase 3 consolidation created new files that inherited logic from multiple sources — those were not explicitly called out in the plan for testing.

### 3. Null Safety in Consolidated Services

All checked consolidated services handle null/undefined correctly:

- **`mpv-control-service.ts`**: Accepts `MpvRuntimeClientLike | null`, uses null checks and optional chaining before all access.
- **`overlay-bridge-service.ts`**: Returns `false` early if window is null or destroyed (`line 17: if (!options.mainWindow || options.mainWindow.isDestroyed()) return false`).
- **`startup-service.ts`**: Uses dependency injection with all deps provided upfront, async operations awaited in sequence, config access uses optional chaining.

**No null safety issues found.**

### 4. Import Consistency

All 68 imports in `main.ts` from `./core/services` resolve to existing exports. No references to deleted files remain. The barrel export in `index.ts` has 79 entries with no dead exports (one minor case: `isGlobalShortcutRegisteredSafe` is only used within the service layer itself, not by main.ts).

---

## Risk Assessment

| Area | Risk | Mitigation |
|------|------|------------|
| Logic regression from inlining | Low | Code moved verbatim, 67 tests pass |
| Overlay rendering regression | Medium | Requires manual smoke test with MPV |
| Unused scaffolding becoming stale | Low | Remove or document with clear intent |
| Missing test coverage on new files | Medium | Add tests for the 3 higher-risk services |
| Race conditions | Low | fieldGroupingResolver race fixed with sequence counter |
| Async error swallowing | Low | All async paths have error boundaries |

---

## Recommendations

1. **Remove unused scaffolding** (`src/core/action-bus.ts`, `src/core/actions.ts`, `src/core/app-context.ts`, `src/core/module-registry.ts`, `src/core/module.ts`, `src/modules/`, `src/ipc/`) — ~500 lines of dead code that contradicts the refactoring goal of reducing unnecessary abstraction.

2. **Add tests for higher-risk consolidated services** — `overlay-shortcut-handler.ts`, `mining-service.ts`, and `anki-jimaku-service.ts` have the most complex logic among the untested new files.

3. **Perform desktop smoke test** — Verify overlay rendering, card mining, and field-grouping interaction in a real session with MPV running.

4. **Consider removing `isGlobalShortcutRegisteredSafe` from barrel export** — It's only used internally by the service layer, not by main.ts.
