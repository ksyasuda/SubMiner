---
id: TASK-20.1
title: Refactor overlay runtime to use per-layer window bounds ownership
status: Done
assignee: []
created_date: '2026-02-12 08:47'
updated_date: '2026-02-13 08:04'
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

- [x] #1 `updateOverlayBoundsService` no longer applies the same bounds to every overlay window by default.
- [x] #2 Main runtime/manager exposes per-layer bounds update paths for visible and invisible overlays.
- [x] #3 Window tracker updates feed shared origin data; each layer applies its own computed bounds.
- [x] #4 Single-layer behavior (visible-only or invisible-only) remains unchanged from user perspective.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Started implementation for per-layer overlay bounds ownership refactor.

Implemented per-layer bounds ownership path: visible and invisible layers now update bounds independently through overlay manager/runtime plumbing, while preserving existing geometry source behavior.

Replaced shared all-window bounds application with per-window bound application service and layer-specific runtime calls from visibility/tracker flows.

Archiving requested by user.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Refactored overlay bounds ownership to per-layer update paths. Tracker geometry remains shared input, but visible/invisible windows apply bounds independently via explicit layer routes. Existing single-layer UX behavior is preserved.

<!-- SECTION:FINAL_SUMMARY:END -->
