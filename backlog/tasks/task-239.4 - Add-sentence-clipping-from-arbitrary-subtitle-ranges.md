---
id: TASK-239.4
title: Add sentence clipping from arbitrary subtitle ranges
status: To Do
assignee: []
created_date: '2026-03-26 20:49'
labels:
  - feature
  - subtitle
  - anki
  - ux
milestone: m-2
dependencies: []
references:
  - src/renderer/modals/subtitle-sidebar.ts
  - src/main/runtime/subtitle-position.ts
  - src/anki-integration/card-creation.ts
  - src/main/runtime/mpv-main-event-actions.ts
parent_task_id: TASK-239
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Current mining flows are optimized around the active subtitle line. Add a sentence-clipping workflow that lets users select an arbitrary contiguous subtitle range, preview the combined text/timing, and mine from that selection. This should improve multi-line dialogue capture without forcing manual copy/paste or separate post-processing.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [ ] #1 Users can select a contiguous subtitle range from the existing subtitle UI instead of being limited to the active cue.
- [ ] #2 The workflow previews the combined text and resulting timing range before mining.
- [ ] #3 Mining from a clipped range uses the combined subtitle payload in card generation while preserving existing single-line behavior.
- [ ] #4 The feature handles overlapping/edge timing cases predictably and does not corrupt the normal active-cue flow.
- [ ] #5 Tests cover range selection, combined payload generation, and at least one card-creation path using a clipped selection.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Define a selection model that fits the existing subtitle sidebar/runtime data flow.
2. Add preview + confirmation UI before routing the clipped payload into mining.
3. Keep the existing single-line path intact and treat clipping as an additive workflow.
4. Verify with subtitle-sidebar, runtime, and Anki/card-creation tests.
<!-- SECTION:PLAN:END -->
