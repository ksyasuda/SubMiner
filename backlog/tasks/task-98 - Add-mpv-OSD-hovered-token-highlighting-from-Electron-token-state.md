---
id: TASK-98
title: Add mpv OSD hovered-token highlighting from Electron token state
status: Done
assignee: []
created_date: '2026-02-21 23:16'
updated_date: '2026-02-22 07:49'
labels:
  - mpv
  - subtitles
  - electron
  - ux
dependencies: []
priority: high
ordinal: 81000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Implement hovered-token highlighting in mpv subtitle OSD by reusing token/hover state that already exists in Electron overlays. When pointer hovers a token in overlay, propagate hovered token identity to mpv subtitle rendering path so the corresponding mpv subtitle token is color-highlighted consistently.
<!-- SECTION:DESCRIPTION:END -->

## Suggestions

<!-- SECTION:SUGGESTIONS:BEGIN -->
- Extend Electron->main IPC hover payload to include stable token index/id usable by mpv renderer.
- Keep mpv highlight style configurable through existing subtitle style settings where possible.
- Reset mpv highlight state on subtitle changes and hover leave events.
<!-- SECTION:SUGGESTIONS:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Identify current overlay hover event source and token identity payload shape.
2. Extend main runtime message path to carry hovered token to mpv OSD renderer service.
3. Update mpv subtitle ASS rendering to apply hover color override to matching token span.
4. Add regression tests for hover enter/move/leave and subtitle replacement edge cases.
5. Verify behavior in plugin flow (`--texthooker` -> `--start`) and normal overlay mode.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Hovering a token in Electron overlay highlights corresponding token in mpv subtitle OSD.
- [x] #2 Hover leave (or subtitle switch) clears mpv hovered-token highlight reliably.
- [x] #3 No regressions in existing token color annotations (known/N+1/frequency/JLPT).
- [x] #4 Automated tests cover hover payload propagation + mpv rendering highlight behavior.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented invisible-overlay-only hovered token highlighting in mpv using a dedicated `osd-overlay` ASS layer. Renderer now tags token spans with stable token indices and reports hover index changes (`subtitle-token-hover:set`) through preload IPC. Main caches latest tokenized subtitle payload, tracks hovered token index, and pushes/clears ASS highlight overlay via new runtime helper `src/main/runtime/mpv-hover-highlight.ts`. Added regression tests for ASS/command generation and IPC dependency wiring. Validation: `bun run build`, `node --test dist/main/runtime/mpv-hover-highlight.test.js dist/core/services/ipc.test.js`, `node --test dist/renderer/subtitle-render.test.js`.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Relevant unit/integration tests pass
- [x] #2 Docs/config notes updated for any new setting or payload contract
<!-- DOD:END -->
