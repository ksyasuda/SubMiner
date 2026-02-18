---
id: TASK-11
title: Break up the applyInvisibleSubtitleLayoutFromMpvMetrics mega function
status: Done
assignee: []
created_date: '2026-02-11 08:21'
updated_date: '2026-02-16 01:34'
labels:
  - refactor
  - renderer
  - complexity
milestone: Codebase Clarity & Composability
dependencies:
  - TASK-27.5
references:
  - src/renderer/renderer.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

In renderer.ts (around lines 865-1075), `applyInvisibleSubtitleLayoutFromMpvMetrics` is a 211-line function with up to 5 levels of nesting. It handles OSD scaling calculations, platform-specific font compensation (macOS vs Linux), DPR calculations, ASS alignment tag interpretation (\an tags), baseline compensation, line-height fixes, font property application, and transform origin — all interleaved.

Extract into focused helpers:

- `calculateOsdScale(metrics, renderAreaHeight)` — pure scaling math
- `calculateSubtitlePosition(metrics, scale, alignment)` — ASS \an tag interpretation + positioning
- `applyPlatformFontCompensation(style, platform)` — macOS kerning/size adjustments
- `applySubtitleStyle(element, computedStyle)` — DOM style application

This can be done independently of or as part of TASK-6 (renderer split).

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 No single function exceeds ~50 lines in the positioning logic
- [x] #2 Helper functions are pure where possible (take inputs, return outputs)
- [x] #3 Platform-specific branches isolated into dedicated helpers
- [x] #4 Invisible overlay positioning still works correctly on Linux and macOS
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Helpers were split so positioning math, base layout, and typography/vertical handling are no longer in one monolith; see `src/renderer/positioning/invisible-layout.ts` and peer files.

Applied as part of TASK-27.5 with helper extraction: moved mpv subtitle layout orchestration to `invisible-layout.ts` and extracted metric/base/style helpers into `invisible-layout-metrics.ts` and `invisible-layout-helpers.ts`.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Decomposition of `applyInvisibleSubtitleLayoutFromMpvMetrics` completed as part of TASK-27.5: function body split into metric/layout/typography helpers and small coordinator preserved. Manual validation completed by user; behavior remains stable.

<!-- SECTION:FINAL_SUMMARY:END -->
