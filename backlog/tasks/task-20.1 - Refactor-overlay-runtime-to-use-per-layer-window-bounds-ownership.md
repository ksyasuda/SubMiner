---
id: TASK-20.1
title: Refactor overlay runtime to use per-layer window bounds ownership
status: To Do
assignee: []
created_date: '2026-02-12 08:47'
updated_date: '2026-02-12 09:42'
labels: []
dependencies: []
parent_task_id: TASK-20
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Refactor overlay runtime so each overlay layer owns and applies its bounds independently. Keep tracker geometry as shared origin input only.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 `updateOverlayBoundsService` no longer applies the same bounds to every overlay window by default.
- [ ] #2 Main runtime/manager exposes per-layer bounds update paths for visible and invisible overlays.
- [ ] #3 Window tracker updates feed shared origin data; each layer applies its own computed bounds.
- [ ] #4 Single-layer behavior (visible-only or invisible-only) remains unchanged from user perspective.
<!-- AC:END -->
