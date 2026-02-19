---
id: TASK-71
title: Split main.ts into domain runtime modules round 2
status: To Do
assignee: []
created_date: '2026-02-18 11:35'
updated_date: '2026-02-18 11:35'
labels:
  - architecture
  - refactor
  - maintainability
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`src/main.ts` remains >3k LOC and still mixes composition, runtime state mutation, domain workflows (Jellyfin, AniList), tray/UI setup, and IPC wiring. This creates high cognitive load and broad change blast radius.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Extract cohesive runtime modules:
   - AniList runtime orchestration
   - Jellyfin runtime orchestration
   - tray runtime
   - MPV event binding/runtime hooks
2. Keep `main.ts` as composition root only (state wiring + module registration).
3. Define narrow service/deps interfaces per module to reduce closure capture.
4. Add/adjust unit tests for extracted modules.
5. Preserve existing CLI/IPC behavior exactly.
6. Update architecture docs with new ownership boundaries.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 `src/main.ts` responsibilities reduced to composition/wiring concerns
- [ ] #2 Extracted runtime modules have focused interfaces and isolated tests
- [ ] #3 No CLI/IPC regressions in existing test suite
- [ ] #4 Docs reflect new module boundaries
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 `bun run test:fast` passes
- [ ] #2 Architecture docs updated
<!-- DOD:END -->

