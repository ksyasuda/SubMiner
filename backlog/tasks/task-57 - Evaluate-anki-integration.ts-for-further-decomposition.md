---
id: TASK-57
title: Evaluate anki-integration.ts for further decomposition
status: To Do
assignee: []
created_date: '2026-02-16 04:47'
updated_date: '2026-02-22 07:14'
labels:
  - maintainability
  - anki
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/anki-integration.ts
  - /home/sudacode/projects/japanese/SubMiner/src/anki-integration/
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

`src/anki-integration.ts` remains a major hotspot after refactors and still centralizes too many adapter and helper responsibilities.

Current state:

- Main file: `src/anki-integration.ts` (~1120 lines on 2026-02-22)
- Extracted modules:
  - workflow and service collaborators already exist under `src/anki-integration/*`
  - main class still owns media generation wrappers, field parsing helpers, config normalization, and notification plumbing

This task decides and documents the next decomposition move, then prepares an implementation-ready extraction plan.

<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Re-map `src/anki-integration.ts` by concern: orchestration, media generation wrappers, field parsing helpers, notification/UI feedback, runtime config patching.
2. Identify cohesive extraction candidates that reduce class surface without adding indirection noise.
3. Define target module boundaries and public interfaces for each candidate extraction.
4. Produce a concrete follow-up implementation plan (candidate files, migration order, regression tests).
5. If no extraction is justified, record explicit rationale and closure criteria.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Current anki-integration responsibilities are documented with concrete line/section mapping.
- [ ] #2 At least one recommended path is chosen: extraction plan or explicit keep-as-is rationale.
- [ ] #3 Decision includes maintainability tradeoffs and expected impact on testability/readability.
- [ ] #4 Follow-up implementation scope (files/tests/sequence) is explicit enough to execute directly.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Decision + plan/rationale is captured in task notes.
- [ ] #2 Any proposed extraction references existing collaborators under `src/anki-integration/*`.
- [ ] #3 Follow-up implementation task(s) are created if extraction is recommended.
<!-- DOD:END -->
