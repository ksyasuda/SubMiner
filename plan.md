# SubMiner Refactoring Plan

## Goals

- Eliminate unnecessary abstraction layers and thin wrapper services
- Consolidate related services into cohesive domain modules
- Fix known bugs and code quality issues
- Add test coverage for critical untested services
- Keep main.ts lean without pushing complexity into pointless indirection

## Guiding Principles

- **Verify after every phase**: `pnpm run build && pnpm run test:config && pnpm run test:core` must pass
- **One concern per commit**: each commit should be a single logical change (can be multiple files as long as it makes sense logically)
- **Inline first, restructure second**: delete the wrapper, verify nothing breaks, then clean up
- **Don't create new abstractions**: the goal is fewer files, not different files

---

## Phase 1: Delete Thin Wrappers (9 files + 7 test files)

These services wrap a single function call or trivial operation. Inline their logic
into callers (mostly main.ts or the service that calls them) and delete both the
service file and its test file.

### 1.1 Inline `config-warning-runtime-service.ts` (14 lines)

Two pure string-formatting functions. Inline the format string directly where
`logConfigWarningRuntimeService` is called (should be in main.ts startup or
`app-logging-runtime-service.ts`). Delete both files.

- Delete: `config-warning-runtime-service.ts`, `config-warning-runtime-service.test.ts`

### 1.2 Inline `overlay-modal-restore-service.ts` (18 lines)

Wraps `Set.add()` and a conditional `Set.delete()`. Inline the Set operations at
call sites in main.ts.

- Delete: `overlay-modal-restore-service.ts`, `overlay-modal-restore-service.test.ts`

### 1.3 Inline `runtime-options-manager-runtime-service.ts` (17 lines)

Wraps `new RuntimeOptionsManager(...)`. Call the constructor directly.

- Delete: `runtime-options-manager-runtime-service.ts`, `runtime-options-manager-runtime-service.test.ts`

### 1.4 Inline `app-logging-runtime-service.ts` (28 lines)

Creates an object with two methods. After inlining config-warning (1.1), this
becomes a trivial object literal. Inline where the logging runtime is created.

- Delete: `app-logging-runtime-service.ts`, `app-logging-runtime-service.test.ts`

### 1.5 Inline `overlay-send-service.ts` (26 lines)

Wraps `window.webContents.send()` with a null/destroyed guard. This is a
one-liner pattern. Inline at call sites.

- Delete: `overlay-send-service.ts` (no test file)

### 1.6 Inline `startup-resource-runtime-service.ts` (26 lines)

Two functions that call a constructor and a check method. Inline into the
app-ready startup flow.

- Delete: `startup-resource-runtime-service.ts`, `startup-resource-runtime-service.test.ts`

### 1.7 Inline `config-generation-runtime-service.ts` (26 lines)

Simple conditional (if args say generate config, call generator, quit). Inline
into the startup bootstrap flow.

- Delete: `config-generation-runtime-service.ts`, `config-generation-runtime-service.test.ts`

### 1.8 Inline `app-shutdown-runtime-service.ts` (27 lines)

Calls 10 cleanup functions in sequence. Inline into the willQuit handler in
main.ts.

- Delete: `app-shutdown-runtime-service.ts`, `app-shutdown-runtime-service.test.ts`

### 1.9 Inline `shortcut-ui-deps-runtime-service.ts` (24 lines)

Single wrapper that unwraps 4 getters and calls `runOverlayShortcutLocalFallback`.
Inline at call site.

- Delete: `shortcut-ui-deps-runtime-service.ts`, `shortcut-ui-deps-runtime-service.test.ts`

### Phase 1 verification

```bash
pnpm run build && pnpm run test:core
```

**Expected result**: 16 files deleted, ~260 lines of service code removed,
~350 lines of test code for trivial wrappers removed. `index.ts` barrel export
shrinks by ~10 entries.

---

## Phase 2: Consolidate DI Adapter Services (4 files)

These are 50-130 line files that map one interface shape to another with minimal
logic. They exist because the service they adapt has a different interface than
what main.ts provides. The fix is to align the interfaces so the adapter isn't
needed, or absorb the adapter into the service it feeds.

### 2.1 Merge `cli-command-deps-runtime-service.ts` into `cli-command-service.ts`

The deps adapter (132 lines) maps main.ts state into `CliCommandServiceDeps`.
Instead: make `handleCliCommandService` accept the same shape main.ts naturally
provides, or accept a smaller interface with the actual values rather than
getter/setter pairs. Move any null-guarding logic into the command handlers
themselves.

- Delete: `cli-command-deps-runtime-service.ts`, `cli-command-deps-runtime-service.test.ts`
- Modify: `cli-command-service.ts` to accept deps directly

### 2.2 Merge `ipc-deps-runtime-service.ts` into `ipc-service.ts`

Same pattern as 2.1. The deps adapter (100 lines) maps main.ts state for IPC
handlers. Merge the defensive null checks and window guards into the IPC handlers
that need them.

- Delete: `ipc-deps-runtime-service.ts`, `ipc-deps-runtime-service.test.ts`
- Modify: `ipc-service.ts`

### 2.3 Merge `tokenizer-deps-runtime-service.ts` into `tokenizer-service.ts`

The adapter (45 lines) has one non-trivial function (`tokenizeWithMecab` with null
checks and token merging). Move that logic into `tokenizer-service.ts`.

- Delete: `tokenizer-deps-runtime-service.ts`, `tokenizer-deps-runtime-service.test.ts`
- Modify: `tokenizer-service.ts`

### 2.4 Merge `app-lifecycle-deps-runtime-service.ts` into `app-lifecycle-service.ts`

The adapter (57 lines) wraps Electron app events. Merge event binding into the
lifecycle service itself since it already knows about Electron's app lifecycle.

- Delete: `app-lifecycle-deps-runtime-service.ts`, `app-lifecycle-deps-runtime-service.test.ts`
- Modify: `app-lifecycle-service.ts`

### Phase 2 verification

```bash
pnpm run build && pnpm run test:core
```

**Expected result**: 8 more files deleted. DI adapters absorbed into the services
they feed. `index.ts` shrinks further.

---

## Phase 3: Consolidate Related Services

Merge services that are split across multiple files but represent a single concern.

### 3.1 Merge overlay visibility files

`overlay-visibility-runtime-service.ts` (46 lines) is a thin orchestration layer
over `overlay-visibility-service.ts` (183 lines). Merge into one file:
`overlay-visibility-service.ts`.

- Delete: `overlay-visibility-runtime-service.ts`
- Modify: `overlay-visibility-service.ts` — absorb the 3 exported functions

### 3.2 Merge overlay broadcast files

`overlay-broadcast-runtime-service.ts` (45 lines) contains utility functions for
window filtering and broadcasting. These are closely related to
`overlay-manager-service.ts` (49 lines) which manages the window references.
Merge broadcast functions into the overlay manager since it already owns the
window state.

- Delete: `overlay-broadcast-runtime-service.ts`, `overlay-broadcast-runtime-service.test.ts`
- Modify: `overlay-manager-service.ts` — add broadcast methods

### 3.3 Merge overlay shortcut files

There are 4 shortcut-related files:

- `overlay-shortcut-service.ts` (169 lines) — registration
- `overlay-shortcut-runtime-service.ts` (105 lines) — runtime handlers
- `overlay-shortcut-lifecycle-service.ts` (52 lines) — sync/refresh/unregister
- `overlay-shortcut-fallback-runner.ts` (114 lines) — local fallback execution

Consolidate into 2 files:

- `overlay-shortcut-service.ts` — registration + lifecycle (absorb lifecycle-service)
- `overlay-shortcut-handler.ts` — runtime handlers + fallback runner

- Delete: `overlay-shortcut-lifecycle-service.ts`, `overlay-shortcut-fallback-runner.ts`

### 3.4 Merge numeric shortcut files

`numeric-shortcut-runtime-service.ts` (37 lines) and
`numeric-shortcut-session-service.ts` (99 lines) are two halves of one feature.
Merge into `numeric-shortcut-service.ts`.

- Delete: `numeric-shortcut-runtime-service.ts`, `numeric-shortcut-runtime-service.test.ts`
- Modify: `numeric-shortcut-session-service.ts` → rename to `numeric-shortcut-service.ts`

### 3.5 Merge startup/bootstrap files

`startup-bootstrap-runtime-service.ts` (53 lines) and
`app-ready-runtime-service.ts` (77 lines) are both startup orchestration.
Merge into a single `startup-service.ts`.

- Delete: `startup-bootstrap-runtime-service.ts`, `startup-bootstrap-runtime-service.test.ts`
- Modify: `app-ready-runtime-service.ts` → rename to `startup-service.ts`, absorb bootstrap

### Phase 3 verification

```bash
pnpm run build && pnpm run test:core
```

**Expected result**: ~10 more files deleted. Related services consolidated into
single cohesive modules.

---

## Phase 4: Fix Bugs and Code Quality

### 4.1 Remove debug `console.log` statements

`overlay-visibility-service.ts` has 7+ raw `console.log`/`console.warn` calls
used for debugging. Remove them or replace with the app logger if the messages
have ongoing diagnostic value.

### 4.2 Fix `as never` type cast

`tokenizer-deps-runtime-service.ts` (or its successor after Phase 2) uses
`return mergeTokens(rawTokens as never)`. Investigate the actual type mismatch
and fix it properly.

### 4.3 Guard `fieldGroupingResolver` against race conditions

In main.ts, the `fieldGroupingResolver` is a single global variable. If two
field grouping requests arrive concurrently, the second overwrites the first's
resolver. Add a request ID or sequence number so stale resolutions are ignored.

### 4.4 Audit async callbacks in CLI command handlers

Verify that async functions passed as callbacks in the CLI command and IPC handler
wiring (main.ts lines ~697-707, ~1347, ~1360) are properly awaited or have
`.catch()` handlers so rejections aren't silently swallowed.

### 4.5 Drop the `-runtime-service` naming convention

After phases 1-3, rename remaining files to just `*-service.ts`. The "runtime"
prefix adds no meaning. Do this as a batch rename commit with no logic changes.

### Phase 4 verification

```bash
pnpm run build && pnpm run test:core
```

Manually smoke test: launch SubMiner, verify overlay renders, mine a card, toggle
field grouping.

---

## Phase 5: Add Tests for Critical Untested Services

These are the highest-risk untested modules. Add focused tests that verify
real behavior, not just that mocks were called.

### 5.1 `mpv-service.ts` (761 lines, untested)

Test the IPC protocol layer:

- Socket message parsing (JSON line protocol)
- Property change event dispatch
- Request/response matching via request IDs
- Reconnection behavior on socket close
- Subtitle text extraction from property changes

### 5.2 `subsync-service.ts` (427 lines, untested)

Test:

- Config resolution (alass vs ffsubsync path selection)
- Command construction for each sync engine
- Timeout and error handling for child processes
- Result parsing

### 5.3 `tokenizer-service.ts` (305 lines, untested)

Test:

- Yomitan parser initialization and readiness
- Token extraction from parsed results
- Fallback behavior when parser unavailable
- Edge cases: empty text, CJK-only, mixed content

### 5.4 `cli-command-service.ts` (204 lines, partially tested)

Expand existing tests to cover:

- All CLI command dispatch paths
- Error propagation from async handlers
- Second-instance argument forwarding

### Phase 5 verification

```bash
pnpm run test:core
```

All new tests pass. Aim for the 4 services above to each have 5-10 meaningful
test cases.

---

## Phase 6 (Optional): Domain-Based Directory Structure

After phases 1-5, the service count should be roughly 20-25 files down from 47.
If that still feels too flat, group by domain:

```
src/core/
  mpv/
    mpv-service.ts
    mpv-render-metrics-service.ts
  overlay/
    overlay-manager-service.ts
    overlay-visibility-service.ts
    overlay-window-service.ts
    overlay-shortcut-service.ts
    overlay-shortcut-handler.ts
  mining/
    mining-runtime-service.ts
    field-grouping-service.ts
    field-grouping-overlay-service.ts
  startup/
    startup-service.ts
    app-lifecycle-service.ts
  ipc/
    ipc-service.ts
    ipc-command-service.ts
  shortcuts/
    shortcut-service.ts
    numeric-shortcut-service.ts
  services/
    tokenizer-service.ts
    subtitle-position-service.ts
    subtitle-ws-service.ts
    texthooker-service.ts
    yomitan-extension-loader-service.ts
    yomitan-settings-service.ts
    secondary-subtitle-service.ts
    runtime-config-service.ts
    runtime-options-runtime-service.ts
    cli-command-service.ts
```

Only do this if the flat directory still feels unwieldy after consolidation.
This is cosmetic and low-priority relative to phases 1-5.

---

## Summary

| Phase                      | Files Deleted       | Files Modified    | Risk   | Effort |
| -------------------------- | ------------------- | ----------------- | ------ | ------ |
| 1. Delete thin wrappers    | 16 (9 svc + 7 test) | main.ts, index.ts | Low    | Small  |
| 2. Consolidate DI adapters | 8 (4 svc + 4 test)  | 4 services        | Medium | Medium |
| 3. Merge related services  | ~10                 | ~5 services       | Medium | Medium |
| 4. Fix bugs & rename       | 0                   | ~6 files          | Low    | Small  |
| 5. Add critical tests      | 0                   | 4 new test files  | Low    | Medium |
| 6. Directory restructure   | 0                   | All imports       | Low    | Small  |

**Net result**: ~34 files deleted, service count from 47 → ~22, index.ts from
92 exports → ~45, and the remaining services each have a clear reason to exist.
