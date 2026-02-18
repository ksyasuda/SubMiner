---
id: TASK-27.4
title: 'Split mpv-service.ts into protocol, transport, property, and facade layers'
status: Done
assignee:
  - backend
created_date: '2026-02-13 17:13'
updated_date: '2026-02-18 04:11'
labels:
  - 'owner:backend'
dependencies:
  - TASK-27.1
references:
  - src/core/services/mpv-service.ts
documentation:
  - docs/architecture.md
parent_task_id: TASK-27
priority: high
ordinal: 41000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Split mpv-service.ts (773 LOC) into thin, testable layers without changing wire protocol behavior.

**This task absorbs the scope of TASK-8** (Reduce MpvIpcClient deps interface and separate protocol from application logic). The work proceeds in two phases:

### Phase A (was TASK-8): Separate protocol from application logic
- Reduce the 22-property `MpvIpcClientDeps` interface to protocol-level concerns only
- Move application-level reactions (subtitle broadcast, overlay visibility sync, timing tracking) to event emitter or external listener pattern
- Make MpvIpcClient testable without mocking 22 callbacks

### Phase B: Physical file split
- Create submodules for: protocol parsing/dispatch, connection lifecycle/retry, property subscriptions/state mapping
- Wire the public facade to maintain API compatibility for all existing consumers
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 MpvIpcClient deps interface reduced to protocol-level concerns only (from current 22 properties).
- [x] #2 Application-level reactions (subtitle broadcast, overlay sync, timing) handled via event emitter or external listeners registered by main.ts.
- [x] #3 Create submodules for protocol parsing/dispatch, connection lifecycle/retry, and property subscriptions/state mapping.
- [x] #4 The public mpv service API (MpvClient interface) remains compatible for all existing consumers.
- [x] #5 MpvIpcClient is testable without mocking 22 callbacks — protocol layer tests need only socket-level mocks.
- [x] #6 Add at least one focused regression check for reconnect + property update flow.
- [x] #7 Document expected event flow in docs/structure-roadmap.md.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
## TASK-8 Absorption

TASK-8 and TASK-27.4 overlapped significantly:
- TASK-8: "Separate protocol from application logic" + shrink deps interface
- TASK-27.4: "Split into protocol, transport, property, and facade layers"

TASK-27.4 is TASK-8 + physical file splitting. Running them as separate tasks would mean touching the same file twice with the same goals. Now consolidated as Phase A (interface separation) and Phase B (file split) within a single task.

**TASK-8 should be marked as superseded** once this task begins.

## Dependency Note
Original plan listed TASK-8 as a dependency. Since TASK-8's scope is now absorbed here, the only remaining dependency is TASK-27.1 (inventory/contracts map).

Started prep: reviewed mpv-service coupling and prepared sequence for protocol/application split; no code split performed yet due current focus on keeping 27.2/27.3 sequencing compatible.

Known compatibility constraint: TASK-27.4 should proceed only after main.ts AppState migration is stable and after the app-level overlay/subsync/anki behavior contracts are preserved.

Milestone progress: extracted protocol buffer parsing into `src/core/services/mpv-protocol.ts`; `src/core/services/mpv-service.ts` now uses `splitMpvMessagesFromBuffer` in `processBuffer` and still delegates full message handling to existing handler. This is a small Phase B step toward protocol/dispatch separation.

Protocol extraction completed: full `MpvMessage` handling moved into `src/core/services/mpv-protocol.ts` via `splitMpvMessagesFromBuffer` + `dispatchMpvProtocolMessage`; `MpvIpcClient` now delegates all message parsing/dispatch through `MpvProtocolHandleMessageDeps` and resolves pending requests through `tryResolvePendingRequest`. `main.ts` wiring remains unchanged.

Updated `docs/structure-roadmap.md` expected mpv flow snapshot to reflect protocol parse/dispatch extraction (`splitMpvMessagesFromBuffer` + `dispatchMpvProtocolMessage`) and façade delegation path via `MpvProtocolHandleMessageDeps`.

Progress update: extracted socket connect/data/error/close/send/reconnect scheduling responsibilities into `MpvSocketTransport` (`src/core/services/mpv-transport.ts`) and wired `MpvIpcClient` to delegate connection lifecycle/send through it. Added `MpvSocketTransport` lifecycle tests in `src/core/services/mpv-transport.test.ts` covering connect/send/error/close behavior. Still in-progress on broader architectural refactor and API boundary reduction for `MpvIpcClient` deps beyond this transport split.

Added focused transport lifecycle regression coverage in `src/core/services/mpv-transport.test.ts`: connect/connect-idempotence, lifecycle callback ordering, and `shutdown()` resets connection/socket state. This covers reconnect/edge-case behavior at transport layer as part of criterion #6 toward protocol + lifecycle regression protection.

Added mpv-service unit regression for close lifecycle: `MpvIpcClient onClose resolves outstanding pending requests and triggers reconnect scheduling path via client transport callbacks (`src/core/services/mpv-service.test.ts`). This complements transport-level lifecycle tests for reconnect behavior regression coverage.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Split mpv-service internals into protocol, transport, and property/state-mapping boundaries; reduced MpvIpcClient deps to protocol-level concerns with event-based app reactions in main.ts; added mpv-service/mpv-transport tests for protocol dispatch, reconnect scheduling, and lifecycle regressions; documented expected event flow in docs/structure-roadmap.md.

Added mpv-service reconnect regression test that asserts a reconnect lifecycle replays mpv property bootstrap commands (`secondary-sub-visibility` reset, `observe_property`, and initial `get_property` state fetches) during reconnection.
<!-- SECTION:FINAL_SUMMARY:END -->
