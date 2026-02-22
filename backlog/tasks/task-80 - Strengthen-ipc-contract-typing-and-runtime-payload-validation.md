---
id: TASK-80
title: Strengthen IPC contract typing and runtime payload validation
status: In Progress
assignee:
  - opencode-task80-ipc-contract
created_date: '2026-02-18 11:43'
updated_date: '2026-02-22 00:21'
labels:
  - ipc
  - type-safety
  - reliability
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
IPC handlers still rely on many `unknown` payload casts in main process paths. This task formalizes typed IPC contracts and validates runtime payloads before dispatch to reduce runtime-only failures.
<!-- SECTION:DESCRIPTION:END -->

## Suggestions

<!-- SECTION:SUGGESTIONS:BEGIN -->
- Define canonical channel map (`channel -> request/response/error types`).
- Add boundary validators for untrusted renderer payloads.
- Keep channel registration centralized to avoid drift.
<!-- SECTION:SUGGESTIONS:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Inventory IPC channels and payload shapes in `src/main/ipc-runtime.ts` and registration callsites.
2. Introduce shared IPC type map and typed registration helpers.
3. Add runtime guards/validators at IPC entry points.
4. Remove unsafe casts where typed contracts are introduced.
5. Add negative tests for malformed payloads and expected error responses.
6. Document IPC contract extension process.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 IPC channels are defined in a typed central contract
- [ ] #2 Runtime payload validation exists for externally supplied IPC data
- [ ] #3 Unsafe cast usage in IPC boundary code is materially reduced
- [ ] #4 Malformed payloads are handled gracefully and test-covered
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Plan of record (2026-02-22):
1) Add central typed IPC contract module at `src/shared/ipc/contracts.ts` and migrate `src/core/services/ipc.ts`, `src/core/services/anki-jimaku-ipc.ts`, and `src/preload.ts` from string literals to contract constants/types.
2) Add runtime IPC payload validators at `src/shared/ipc/validators.ts` for externally supplied payloads (runtime option id/direction/value boundary, subsync request shape, overlay modal, subtitle position, and kiku/jimaku payloads where renderer-supplied).
3) Wire validators at IPC boundaries so malformed payloads are handled gracefully (return structured `{ ok: false, error }` for invoke handlers or no-op/log for fire-and-forget channels) and avoid unsafe `as` casts in boundary code.
4) Reduce unsafe casts in runtime IPC wiring (`src/main/dependencies.ts`, `src/main.ts`, IPC composer generics) by narrowing types before domain calls.
5) Add/extend IPC tests for malformed payload behavior (`src/core/services/ipc.test.ts`, `src/core/services/anki-jimaku-ipc.test.ts`), then run `bun run build`, `bun run test:core:src`, and `bun run test:core:dist`.
6) Update `docs/architecture.md` with central IPC contract and boundary-validation conventions; then finalize TASK-80 AC/DoD evidence in Backlog MCP.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-02-22: Started execution session opencode-task80-ipc-contract-20260222T001728Z-obrv. Loading IPC boundary code and preparing implementation plan via writing-plans before any code edits.

Saved plan document: docs/plans/2026-02-22-task-80-ipc-contract-validation.md. Proceeding with executing-plans implementation flow as requested.
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 IPC-related tests pass
- [ ] #2 IPC contract docs updated
<!-- DOD:END -->
