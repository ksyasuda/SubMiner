---
id: TASK-231
title: Restore controller input while subtitle sidebar is open
status: Done
assignee:
  - '@codex'
created_date: '2026-03-24 00:15'
updated_date: '2026-03-24 00:15'
labels:
  - bug
  - controller
  - subtitle-sidebar
  - overlay
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/renderer/renderer.ts
  - /home/sudacode/projects/japanese/SubMiner/src/renderer/controller-interaction-blocking.ts
  - /home/sudacode/projects/japanese/SubMiner/src/renderer/controller-interaction-blocking.test.ts
priority: high
ordinal: 54900
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

When keyboard-only mode is active, opening the subtitle sidebar should not disable controller navigation and lookup/mining controls. Restore controller input while the sidebar is open, while keeping true modal dialogs blocking controller actions.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Opening the subtitle sidebar does not block controller input for keyboard-only mode actions.
- [x] #2 Controller-select/debug and other true modal dialogs still block controller actions while open.
- [x] #3 Focused regression coverage exists for the sidebar-open controller gating rule.

<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Root cause: renderer gamepad polling used the broad `isAnyModalOpen()` check as its interaction gate, and that list includes `subtitleSidebarModalOpen`. The subtitle sidebar is non-modal for controller usage, so gamepad input was being suppressed whenever the sidebar was visible.

Fixed by extracting a dedicated controller-interaction blocking helper that excludes the subtitle sidebar but keeps the existing blocking behavior for true modal dialogs.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Restored controller input while the subtitle sidebar is open by switching gamepad polling to a dedicated modal-blocking rule that leaves the sidebar controller-passive. Added a regression test covering the sidebar-open exception and preserving hard blocks for actual modal dialogs.

<!-- SECTION:FINAL_SUMMARY:END -->
