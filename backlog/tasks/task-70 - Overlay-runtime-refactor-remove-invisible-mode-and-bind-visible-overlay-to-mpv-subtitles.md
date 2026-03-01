---
id: TASK-70
title: >-
  Overlay runtime refactor: remove invisible mode and bind visible overlay to
  mpv subtitles
status: Done
assignee: []
created_date: '2026-02-28 02:38'
updated_date: '2026-02-28 22:36'
labels: []
dependencies: []
references:
  - 'commit:a14c9da'
  - 'commit:74554a3'
  - 'commit:75442a4'
  - 'commit:dde51f8'
  - 'commit:9e4e588'
  - src/main/overlay-runtime.ts
  - src/main/runtime/overlay-mpv-sub-visibility.ts
  - src/renderer/renderer.ts
  - docs/plans/2026-02-26-secondary-subtitles-main-overlay.md
priority: medium
ordinal: 1000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Scope: Branch-only commits main..HEAD on refactor-overlay (a14c9da through 9e4e588) rebuilt overlay behavior around visible overlay mode and removed legacy invisible overlay paths.

Delivered behavior:
- Removed renderer invisible overlay layout/offset helpers and main hover-highlight runtime code paths.
- Added explicit overlay-to-mpv subtitle visibility synchronization so visible overlay state controls primary subtitle visibility consistently.
- Hardened overlay runtime/bootstrap lifecycle around modal fallback open state and bridge send path edge cases.
- Updated plugin/config/docs defaults to reflect visible-overlay-first behavior and subtitle binding controls.

Risk/impact context:
- Large cross-layer refactor touching runtime wiring, renderer event handling, and plugin behavior.
- Regression coverage added/updated for overlay runtime, mpv protocol handling, renderer cleanup, and subtitle rendering paths.
<!-- SECTION:DESCRIPTION:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Completed and validated in branch commit set before merge. Refactor reduces dead overlay modes, centralizes subtitle visibility behavior, and documents new defaults/constraints.
<!-- SECTION:FINAL_SUMMARY:END -->
