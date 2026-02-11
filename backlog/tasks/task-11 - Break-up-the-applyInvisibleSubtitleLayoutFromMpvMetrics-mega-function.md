---
id: TASK-11
title: Break up the applyInvisibleSubtitleLayoutFromMpvMetrics mega function
status: To Do
assignee: []
created_date: '2026-02-11 08:21'
labels:
  - refactor
  - renderer
  - complexity
milestone: Codebase Clarity & Composability
dependencies: []
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
