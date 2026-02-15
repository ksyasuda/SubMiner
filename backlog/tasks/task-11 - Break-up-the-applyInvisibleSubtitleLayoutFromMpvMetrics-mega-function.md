---
id: TASK-11
title: Break up the applyInvisibleSubtitleLayoutFromMpvMetrics mega function
status: To Do
assignee: []
created_date: '2026-02-11 08:21'
updated_date: '2026-02-15 07:00'
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
- [ ] #1 No single function exceeds ~50 lines in the positioning logic
- [ ] #2 Helper functions are pure where possible (take inputs, return outputs)
- [ ] #3 Platform-specific branches isolated into dedicated helpers
- [ ] #4 Invisible overlay positioning still works correctly on Linux and macOS
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Reparented as a dependency of TASK-27.5: the mega-function lives in positioning.ts (513 LOC), which is the exact file TASK-27.5 targets for splitting. Decomposing this function is a natural part of that file split. Should be executed together with TASK-27.5.
<!-- SECTION:NOTES:END -->
