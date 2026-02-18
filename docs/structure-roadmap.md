# Structure Roadmap for TASK-27

Date: 2026-02-14

## 1) Oversized refactor candidates (>=400 LOC)

| File                           | Concern                                           | Status                           | Reason                                                                                                                 |
| ------------------------------ | ------------------------------------------------- | -------------------------------- | ---------------------------------------------------------------------------------------------------------------------- |
| `src/main.ts`                  | Bootstrap / composition root / orchestration      | Active                           | Main entrypoint owns startup, lifecycle orchestration, service construction, state mutation surfaces, and IPC bindings |
| `src/anki-integration.ts`      | Domain service orchestration / integrations       | Active                           | 2.6k+ LOC, high cyclomatic coupling to mpv/subtitle timing and mining flows                                            |
| `src/core/services/mpv.ts`     | MPV protocol + app state bridge                   | Active                           | ~780 LOC, large protocol and behavior mix, 22-entry dep interface                                                      |
| `src/core/services/subsync.ts` | Subsync orchestration (ffsubsync/alass workflows) | Active                           | ~494 LOC, file IO + mpv command orchestration + result shaping                                                         |
| `src/renderer/positioning.ts`  | Renderer positioning layout policy                | Active (downstream of TASK-27.5) | 513 LOC, layout/rules + platform-specific behavior in one module                                                       |
| `src/config/service.ts`        | Config load/validation                            | Active support                   | ~601 LOC, central schema validation + mutation APIs                                                                    |
| `src/types.ts`                 | Shared contract surface                           | Active support                   | ~640 LOC, foundational type exports driving module boundaries                                                          |
| `src/config/definitions.ts`    | Config schema/registry definitions                | Active support                   | ~480 LOC, dense constants/interfaces used by config runtime and docs                                                   |
| `src/media-generator.ts`       | Media generation helpers                          | Active support                   | ~431 LOC, utility-heavy with multiple generation flows                                                                 |

## 2) API contracts by target file

### `src/main.ts` (bootstrap + composition root)

- **Entry points**: Electron main process boot path (executed by package `main` entry).
- **Contract responsibilities**:
  - parse CLI/env and select startup flow through `runStartupBootstrapRuntimeService`
  - register app-level lifecycle and command handlers (`runAppLifecycleService`, `handleCliCommandService`)
  - instantiate core services (mpv, overlay runtime, IPC handlers, shortcuts, tokenizer, config services)
  - hold mutable application runtime state:
    - parser/Yomitan windows and extension state
    - mpv client and tracker/state
    - overlay/app bootstrap flags and modal/shortcut state
    - runtime option and anki integration containers
- **Primary callers/integration points**:
  - Electron event loop (`app.whenReady`, `process` signals)
  - startup/bootstrap services, service adapters in `src/core/services`

### `src/core/services/mpv.ts` (protocol + facade)

- **Core exports**: `MpvIpcClient`, `MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY`, `MpvIpcClientDeps`
- **Primary responsibilities / API**:
  - connection management (`connect`, `setSocketPath`, `request`, `send`, `requestProperty`)
  - protocol message handling (`processBuffer`, private message handlers for `property-change` and request IDs)
  - reconnection lifecycle (`scheduleReconnect`, `failPendingRequests`)
  - mpv control actions (`setSubVisibility`, `replayCurrentSubtitle`, `playNextSubtitle`)
  - state restoration (`restorePreviousSecondarySubVisibility`)
- **Current caller set**:
  - `src/main.ts` (construction + lifecycle + service invocations)
  - `src/main/ipc-mpv-command.ts` (runtime control API)
  - `src/core/services/subsync.ts` (`requestProperty`, request ID usage)
  - tests under `src/core/services/mpv.test.ts`
- **Observed coupling risk**:
  - `MpvIpcClientDeps` mixes protocol config with app-level side effects (subtitle broadcast, tokenizer, overlay updates, config reads)

### `src/anki-integration.ts`

- **Class export**: `AnkiIntegration`
- **Primary public operations (validated from external callsites)**:
  - lifecycle: `start`, `stop`, `destroy`
  - card flow: `createSentenceCard`, `updateLastAddedFromClipboard`, `triggerFieldGroupingForLastAddedCard`, `markLastCardAsAudioCard`
  - runtime patching: `applyRuntimeConfigPatch`
- **State dependencies (constructor)**:
  - config, subtitle timing tracker, mpv client, OSD/notification callbacks
- **Primary callers**:
  - `src/core/services/overlay-runtime-init.ts` (initial integration creation)
  - `src/core/services/anki-jimaku.ts` (enable/disable and field-grouping RPC)
  - `src/core/services/mining.ts` (delegates mining actions)

### `src/core/services/subsync.ts`

- **Exports**: `runSubsyncManualService`, `openSubsyncManualPickerService`, `triggerSubsyncFromConfigService`
- **Caller set**:
  - `src/core/services/subsync-runner.ts` (runtime wrappers)
  - `src/core/services/mpv-jimaku/`? (through runtime services and IPC command handlers)

### `src/config/service.ts`

- **Class export**: `ConfigService`
- **Primary API**:
  - `reloadConfig`, `saveRawConfig`, `patchRawConfig`
  - read methods: `getConfig`, `getRawConfig`, `getWarnings`, `getConfigPath`
- **Primary callers**:
  - `src/main.ts`, `src/core/services/cli-command` and runtime config consumers

### `src/renderer/positioning.ts`

- **Public style/runtime contract**
  - exported calculation helpers for subtitle Y-position and inset offsets
  - event handlers for window geometry and subtitle-size changes
- **Primary callers**:
  - `src/renderer/handlers/*`
  - `src/renderer/subtitle-render.ts`

## 3) Split sequence / dependency ordering

Adopted sequence (from TASK-27 parent):

1. `TASK-27.3` (Anki integration split)
2. `TASK-27.2` (main.ts composition-root split) — depends on TASK-7 and TASK-27.3
3. `TASK-27.4` (mpv-service protocol/transport/property/facade split) — depends on TASK-27.1 and absorbs TASK-8
4. `TASK-27.5` (renderer positioning split) — downstream

## 4) Compatibility and migration risks per split

### TASK-27.3 / anki integration surface

- Risk: interface breakage in field-grouping preview/build/enable flow
- Mitigation: keep constructor params and public methods stable; preserve IPC payload shapes

### TASK-27.2 (composition root)

- Risk: lifecycle/cleanup regressions from moving startup hooks and shutdown behavior
- Mitigation: preserve service construction order and keep existing event registration boundaries

Migration note:

- `src/main.ts` now delegates composition edges to `src/main/startup.ts`, `src/main/app-lifecycle.ts`, `src/main/startup-lifecycle.ts`, `src/main/ipc-runtime.ts`, `src/main/cli-runtime.ts`, and `src/main/subsync-runtime.ts`.
- Overlay/modal interaction has been moved into `src/main/overlay-runtime.ts` (window selection, modal restore set tracking, runtime-options palette/modal close handling) so `main.ts` now uses a dedicated runtime service for modal routing instead of inline window bookkeeping.
- Stateful runtime session data has moved to `src/main/state.ts` via `createAppState()` so `main.ts` no longer owns the `AppState` shape inline, only importing and mutating the shared state instance.
- Behavioral contract remains stable: startup flow, CLI dispatch, IPC handlers, and subsync orchestration keep existing external behavior; only dependency wiring moved out of `main.ts`.

### TASK-27.4 (mpv-service)

- Risk: request/deps interface changes could break control and subsync interactions
- Mitigation: preserve public `MpvClient` methods, request semantics, and reconnect events while splitting internal modules

### TASK-27.4 expected event flow snapshot (current)

- `connect()` establishes socket and starts `observe_property` subscriptions via `subscribeToProperties()`.
- `processBuffer()` uses `splitMpvMessagesFromBuffer()` to parse JSON lines and route each message to `handleMessage()`.
- `dispatchMpvProtocolMessage()` now owns protocol-level handling for:
  - `event === "property-change"` updates (subtitle text/ASS/timing, media/path/title, track metrics, secondary subtitles, visibility restore flags)
  - `request_id` responses for startup state fetches and dynamic property queries
  - shutdown handling and pending request resolution
- `handleMessage()` now delegates state mutation and side effects through `MpvProtocolHandleMessageDeps` to the client facade (`emit(...)`, state fields, `sendCommand`, etc.).
- Reconnect path stays behavior-critical:
  - socket close/error clears pending requests and calls `scheduleReconnect()`
  - timer callback calls `connect()` after exponential-ish delay

### TASK-27.5 (renderer positioning)

- Risk: UI placement drift on platform edge cases
- Mitigation: keep existing DOM state updates and geometry math in place while refactoring module boundaries

## 5) Global smoke checklist (required for all TASK-27 subtasks)

- App starts and connects to MPV
- Subtitle text appears in overlay
- Card mining creates a note in Anki
- Field grouping modal opens and resolves
- Global shortcuts work (mine, toggle overlay, copy subtitle)
- Secondary subtitle display works
- TypeScript compiles with no new errors
- Manual smoke checklist executed for runtime behavior
