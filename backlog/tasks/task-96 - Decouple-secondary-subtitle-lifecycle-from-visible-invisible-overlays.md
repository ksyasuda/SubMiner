---
id: TASK-96
title: Decouple secondary subtitle lifecycle from visible/invisible overlays
status: To Do
assignee: []
created_date: '2026-02-21 04:41'
updated_date: '2026-02-21 04:41'
labels:
  - subtitles
  - overlay
  - architecture
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Secondary subtitle behavior should not depend on visible/invisible overlay state transitions. Introduce an independent lifecycle so secondary subtitle rendering, visibility mode (`always`/`hover`/`never`), and positioning stay stable even when primary overlays are toggled or rebound.
<!-- SECTION:DESCRIPTION:END -->

## Suggestions

<!-- SECTION:SUGGESTIONS:BEGIN -->
- Isolate secondary subtitle state management from primary overlay window orchestration.
- Route secondary subtitle updates through a dedicated service/controller boundary.
- Keep MPV secondary subtitle property handling independent from overlay visibility toggles.
<!-- SECTION:SUGGESTIONS:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Inventory existing coupling points between secondary subtitle updates and overlay visibility/bounds services.
2. Introduce explicit secondary subtitle lifecycle state and transitions.
3. Refactor event wiring so visible/invisible overlay toggles do not mutate secondary subtitle state.
4. Validate display modes (`always`/`hover`/`never`) continue to work with independent lifecycle.
5. Add regression tests for overlay toggles, reconnect/restart, and mode-switch behavior.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Toggling visible or invisible overlays does not alter secondary subtitle lifecycle state.
- [ ] #2 Secondary subtitle display mode behavior remains correct across overlay state transitions.
- [ ] #3 Secondary subtitle behavior survives MPV reconnect/restart without overlay-coupling regressions.
- [ ] #4 Automated tests cover decoupled lifecycle behavior and prevent re-coupling.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Relevant unit/integration tests pass
- [ ] #2 Documentation/comments updated where lifecycle ownership changed
<!-- DOD:END -->
