---
id: TASK-30.3
title: Expose resolver operations via Electron IPC to renderer
status: To Do
assignee: []
created_date: '2026-02-13 18:32'
updated_date: '2026-02-13 18:34'
labels: []
dependencies:
  - TASK-30.2
parent_task_id: TASK-30
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add a typed preload and main-IPC contract for streaming queries and playback resolution so the renderer can initiate search/list/resolve without embedding network/provider logic in UI code.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Define IPC handlers in main with input/output schema validation and timeouts.
- [ ] #2 Expose corresponding functions in preload `window.electronAPI` and ElectronAPI types.
- [ ] #3 Reuse existing mpv command channel for playback and add a dedicated request/response flow for resolver actions.
- [ ] #4 Implement safe serialization and error marshalling for resolver-specific failures.
- [ ] #5 Add runtime wiring and lifetime management in app startup/shutdown.
- [ ] #6 Document event/callback behavior for loading/error states.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Phase 3 — API surface: IPC/preload contract for resolver operations
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Renderer code can query providers without importing Node-only modules.
- [ ] #2 IPC paths have clear names and consistent response shapes across all calls.
- [ ] #3 Error paths return explicit machine-readable codes mapped to user-visible messages.
<!-- DOD:END -->
