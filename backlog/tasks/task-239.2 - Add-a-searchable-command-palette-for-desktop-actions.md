---
id: TASK-239.2
title: Add a searchable command palette for desktop actions
status: To Do
assignee: []
created_date: '2026-03-26 20:49'
labels:
  - feature
  - ux
  - desktop
  - shortcuts
milestone: m-2
dependencies: []
references:
  - src/renderer
  - src/shared/ipc/contracts.ts
  - src/main/runtime/overlay-runtime-options.ts
  - src/main.ts
parent_task_id: TASK-239
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
SubMiner already exposes many actions through scattered shortcuts, menus, and modal flows. Add a searchable command palette so users can discover and execute high-value desktop actions from one keyboard-first surface. Build on the existing runtime-options/modal infrastructure where practical instead of creating a completely separate interaction model.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [ ] #1 A keyboard-accessible command palette opens from the desktop app and lists supported actions with searchable labels.
- [ ] #2 Commands are backed by an explicit registry so action availability and labels are not hard-coded in one renderer component.
- [ ] #3 Users can navigate and execute commands entirely from the keyboard.
- [ ] #4 The first slice includes the highest-value existing actions rather than trying to cover every possible command on day one.
- [ ] #5 Tests cover command filtering, execution dispatch, and at least one disabled/unavailable command state.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Define a small command-registry contract shared across renderer and main-process dispatch.
2. Reuse existing modal/runtime plumbing where it fits so the palette is a thin discoverability layer over current actions.
3. Ship a narrow but useful initial command set, then expand later based on usage.
4. Verify with renderer tests plus targeted IPC/runtime tests.
<!-- SECTION:PLAN:END -->
