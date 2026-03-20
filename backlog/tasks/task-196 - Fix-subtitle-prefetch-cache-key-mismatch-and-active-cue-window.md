---
id: TASK-196
title: Fix subtitle prefetch cache-key mismatch and active-cue window
status: Done
assignee: []
created_date: '2026-03-18 16:05'
labels: []
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/core/services/subtitle-processing-controller.ts
  - /home/sudacode/projects/japanese/SubMiner/src/core/services/subtitle-prefetch.ts
documentation: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Investigate and fix file-backed subtitle annotation latency where prefetch should warm upcoming lines but live playback still tokenizes each subtitle line. Likely causes: cache-key mismatch between parsed cue text and mpv `sub-text`, and priority-window selection skipping the currently active cue during mid-line starts/seeks.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Prefetched subtitle entries are reused when live subtitle text differs only by normalization details such as ASS `\N`, newline collapsing, or surrounding whitespace.
- [x] #2 Priority-window selection includes the currently active cue when playback starts or seeks into the middle of a cue.
- [x] #3 Regression tests cover the cache-hit normalization path and active-cue priority-window behavior.
- [x] #4 Verification covers the touched prefetch/controller lane.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing regression tests in `subtitle-processing-controller.test.ts` and `subtitle-prefetch.test.ts`.
2. Normalize cache keys in the subtitle processing controller so prefetch/live paths share keys.
3. Adjust prefetch priority-window selection to include the active cue.
4. Run targeted tests, then SubMiner verification lane for touched files.
<!-- SECTION:PLAN:END -->

## Outcome

<!-- SECTION:OUTCOME:BEGIN -->
Normalized subtitle cache keys inside the processing controller so prefetched ASS/VTT/live subtitle text variants reuse the same cache entry, and changed priority-window selection to include the currently active cue based on cue end time. Added regression coverage for both paths and verified the change with the `core` lane.
<!-- SECTION:OUTCOME:END -->
