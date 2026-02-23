---
id: TASK-110
title: Split overlay into top secondary bar and bottom primary region
status: Done
assignee:
  - codex
created_date: '2026-02-23 02:16'
updated_date: '2026-02-23 02:29'
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

<!-- SECTION:NOTES:END -->
