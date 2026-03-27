---
id: TASK-239
title: Mining workflow upgrades: prioritize high-value user-facing improvements
status: To Do
assignee: []
created_date: '2026-03-26 20:49'
labels:
  - feature
  - ux
  - planning
milestone: m-2
dependencies: []
references:
  - src/main.ts
  - src/renderer
  - src/anki-integration.ts
  - src/config/service.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Track the next set of high-value workflow improvements surfaced by the March 2026 review. The goal is to capture bounded, implementation-sized feature slices with clear user value and avoid prematurely committing to much larger bets like hard-sub OCR, plugin marketplace infrastructure, or cloud config sync. Focus this parent task on features that improve the core mining workflow directly: profile-aware setup, action discoverability, previewing output before mining, and selecting richer subtitle ranges.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [ ] #1 Child tasks exist for the selected near-to-medium-term workflow upgrades with explicit scope and exclusions.
- [ ] #2 The parent task records the recommended sequencing so future work starts with the best value/risk ratio.
- [ ] #3 The tracked feature set stays grounded in existing product surfaces instead of speculative external-platform integrations.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Recommended sequencing:

1. Start TASK-239.3 first. Template preview is the smallest high-signal UX win on a core mining path.
2. Start TASK-239.2 next. A command palette improves discoverability across existing actions without large backend upheaval.
3. Start TASK-239.4 after the preview/palette work. Sentence clipping is high-value but touches runtime, subtitle selection, and card creation flows together.
4. Keep TASK-239.1 as a foundation project and scope it narrowly to local multi-profile support. Do not expand it into cloud sync in the same slice.

Deliberate exclusions for now:

- hard-sub OCR
- plugin marketplace infrastructure
- cloud/device sync
- site-specific streaming source auto-detection beyond narrow discovery spikes
<!-- SECTION:PLAN:END -->
