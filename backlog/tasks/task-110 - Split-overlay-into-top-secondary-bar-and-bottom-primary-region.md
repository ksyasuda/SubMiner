---
id: TASK-110
title: Split overlay into top secondary bar and bottom primary region
status: Done
assignee:
  - codex
created_date: '2026-02-23 02:16'
updated_date: '2026-02-23 02:53'
labels:
  - overlay
  - subtitle
  - ux
dependencies: []
priority: high
ordinal: 110000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Implement a 3-window overlay layout:

- Dedicated secondary subtitle window anchored to top of mpv bounds (max 20% of height).
- Visible overlay window constrained to remaining lower region.
- Invisible overlay window constrained to remaining lower region.

Secondary subtitle bar must stay independently anchored while visible/invisible overlays swap.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Runtime creates/manages a third overlay window dedicated to secondary subtitles.
- [x] #2 Secondary window stays anchored to top region capped at 20% of tracked mpv bounds.
- [x] #3 Visible and invisible overlay windows are constrained to remaining lower region.
- [x] #4 Secondary subtitle rendering/mode updates reach dedicated top window without duplicate top bars in primary windows.
- [x] #5 Focused runtime/core tests cover geometry split + window wiring regressions.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Closure verification plan (2026-02-23):
1) Re-read TASK-110 via Backlog MCP and confirm Done status + AC/summary/notes completeness.
2) Validate evidence commit `b8f7d5e` still matches shipped scope (overlay secondary top bar + primary lower region split).
3) Apply backlog metadata sync only if any field is stale/missing; keep status Done.
4) Record closure verification note with timestamp and validation outcome.
Plan artifact: `docs/plans/2026-02-23-task-110-overlay-closure-verification.md`.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
- Added new `secondary` overlay window kind and runtime factory wiring, plus manager ownership (`get/setSecondaryWindow`) and bounds setter.
- Added geometry splitter `splitOverlayGeometryForSecondaryBar` (20% top secondary, 80% bottom primary), integrated into `updateVisibleOverlayBounds` / `updateInvisibleOverlayBounds` flow.
- Main runtime now creates secondary window alongside primary overlays and syncs secondary window visibility with `secondarySubMode`.
- Renderer now recognizes `layer=secondary`; secondary bar is hidden on primary layers, and primary subtitle/modals are hidden on secondary layer.
- Secondary layer no longer reports overlay content measurements to avoid invalid payloads.
- Validation:
  - `bun run tsc --noEmit`
  - `bun test src/main/runtime/overlay-window-factory.test.ts src/main/runtime/overlay-window-factory-main-deps.test.ts src/main/runtime/overlay-window-runtime-handlers.test.ts src/renderer/error-recovery.test.ts src/core/services/overlay-window.test.ts`
  - `bun run build`
  - `node --test dist/core/services/overlay-manager.test.js`
- Commit: `b8f7d5e` (`feat(overlay): split secondary subtitles into dedicated top window`)

2026-02-23 verification pass: revalidated TASK-110 closure against Backlog finalization criteria. Confirmed status remains Done, all 5 acceptance criteria are checked, implementation notes/final summary are present, and commit `b8f7d5e` still matches shipped scope (secondary top overlay window, 20%/80% geometry split, primary-layer duplicate prevention, focused runtime/core/renderer tests). No additional code changes required.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented the requested 3-window overlay architecture:

- Secondary subtitles render in a dedicated top overlay window.
- Top secondary bar is capped at 20% of mpv bounds.
- Visible/invisible overlays are constrained to the lower remaining region.
- Secondary mode changes now directly control secondary window visibility.
- Primary overlay layers no longer duplicate secondary subtitle rendering.

All acceptance criteria are complete and merged in commit `b8f7d5e`.
<!-- SECTION:FINAL_SUMMARY:END -->
