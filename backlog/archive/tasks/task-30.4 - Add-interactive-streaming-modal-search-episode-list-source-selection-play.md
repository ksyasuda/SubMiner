---
id: TASK-30.4
title: 'Add interactive streaming modal (search, episode list, source selection, play)'
status: To Do
assignee: []
created_date: '2026-02-13 18:32'
updated_date: '2026-02-13 18:34'
labels: []
dependencies:
  - TASK-30.3
parent_task_id: TASK-30
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Implement a renderer flow to query configured providers, display results, let user choose series and episode, and trigger playback for a selected stream. The UI should support keyboard interactions and surface backend errors clearly.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Create modal UI/state model for query, results list, selected item, episode list, candidate qualities, and loading/error status.
- [ ] #2 Wire renderer actions to new IPC methods for search/episode/resolve.
- [ ] #3 Render one-click or enter-to-play action that calls existing mpv playback pathway.
- [ ] #4 Persist minimal user preference (last provider/quality where possible) for faster repeat use.
- [ ] #5 Provide empty/error states and accessibility-friendly focus/keyboard navigation for lists.
- [ ] #6 Add a no-network mode fallback message when resolver calls fail.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Phase 4 — UX: interactive modal flow and playback callout

<!-- SECTION:NOTES:END -->

## Definition of Done

<!-- DOD:BEGIN -->

- [ ] #1 Modal state is isolated and unsubscribes listeners on close.
- [ ] #2 No direct network logic in renderer beyond IPC calls.
- [ ] #3 Visual style and behavior are consistent with existing modal patterns.
<!-- DOD:END -->
