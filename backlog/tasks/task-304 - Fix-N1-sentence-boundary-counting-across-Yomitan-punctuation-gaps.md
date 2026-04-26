---
id: TASK-304
title: Fix N+1 sentence boundary counting across Yomitan punctuation gaps
status: In Progress
assignee: []
created_date: '2026-04-26 05:33'
labels:
  - bug
  - tokenizer
  - annotations
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
N+1 target selection should respect sentence-ending punctuation from the original subtitle text even when Yomitan token output omits punctuation tokens. Current behavior can treat multiple subtitle sentences as one token span and incorrectly satisfy the minimum content-token threshold.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 A subtitle like `てんめ！ふざけんなよ！` does not mark `ふざけん`/similar single-content-token second sentence as N+1 when the minimum sentence word count is 3.
- [ ] #2 N+1 sentence segmentation uses original subtitle text offsets or equivalent source-boundary data, not only punctuation tokens returned by Yomitan.
- [ ] #3 Existing annotation exclusion behavior for particles/grammar tokens remains unchanged.
- [ ] #4 Regression tests cover Yomitan-style token streams where punctuation is absent from the token list.
<!-- AC:END -->
