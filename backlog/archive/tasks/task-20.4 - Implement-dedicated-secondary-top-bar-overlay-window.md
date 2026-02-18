---
id: TASK-20.4
title: Implement dedicated secondary top-bar overlay window
status: To Do
assignee: []
created_date: '2026-02-12 09:43'
labels: []
dependencies: []
parent_task_id: TASK-20
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Create and integrate a dedicated secondary subtitle overlay window with independent lifecycle, z-order, bounds, and pointer policy, decoupled from primary visible/invisible overlay windows.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 A third overlay window dedicated to secondary subtitles is created and managed alongside existing visible/invisible windows.
- [ ] #2 Secondary window visibility follows secondary mode semantics (`hidden`/`visible`/`hover`) independent of primary overlay visibility.
- [ ] #3 Secondary subtitle text/mode/style updates are routed directly to the secondary window renderer path.
- [ ] #4 Pointer passthrough/interaction behavior for secondary window is explicit and does not regress existing hover/selection interactions.
- [ ] #5 Window cleanup/lifecycle (create, close, restore) integrates with existing overlay runtime lifecycle.
<!-- AC:END -->
