---
id: TASK-80
title: Strengthen IPC contract typing and runtime payload validation
status: To Do
assignee: []
created_date: '2026-02-18 11:43'
updated_date: '2026-02-18 11:43'
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

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 IPC-related tests pass
- [ ] #2 IPC contract docs updated
<!-- DOD:END -->

