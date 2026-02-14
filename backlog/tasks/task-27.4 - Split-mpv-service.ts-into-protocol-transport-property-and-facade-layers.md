---
id: TASK-27.4
title: 'Split mpv-service.ts into protocol, transport, property, and facade layers'
status: To Do
assignee:
  - backend
created_date: '2026-02-13 17:13'
updated_date: '2026-02-13 21:15'
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
- [ ] #1 MpvIpcClient deps interface reduced to protocol-level concerns only (from current 22 properties).
- [ ] #2 Application-level reactions (subtitle broadcast, overlay sync, timing) handled via event emitter or external listeners registered by main.ts.
- [ ] #3 Create submodules for protocol parsing/dispatch, connection lifecycle/retry, and property subscriptions/state mapping.
- [ ] #4 The public mpv service API (MpvClient interface) remains compatible for all existing consumers.
- [ ] #5 MpvIpcClient is testable without mocking 22 callbacks — protocol layer tests need only socket-level mocks.
- [ ] #6 Add at least one focused regression check for reconnect + property update flow.
- [ ] #7 Document expected event flow in docs/structure-roadmap.md.
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
<!-- SECTION:NOTES:END -->
