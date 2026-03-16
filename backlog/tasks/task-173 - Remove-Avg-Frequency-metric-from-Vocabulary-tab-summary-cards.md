---
id: TASK-173
title: Remove Avg Frequency metric from Vocabulary tab summary cards
status: Done
assignee: []
created_date: '2026-03-15 00:13'
updated_date: '2026-03-16 05:13'
labels:
  - stats
  - ui
dependencies: []
priority: low
ordinal: 17500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
User requested removing the Avg Frequency card/metric because it is not useful. Remove the UI card and stop computing/storing the summary field in dashboard summary shaping code.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Vocabulary tab no longer renders an "Avg Frequency" stat card.
- [x] #2 Vocabulary summary model no longer exposes or computes averageFrequency.
- [x] #3 Typecheck/tests covering dashboard summary and vocabulary tab pass.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Removed the Vocabulary tab "Avg Frequency" card and deleted the corresponding `averageFrequency` field from `VocabularySummary` and `buildVocabularySummary`.

Verification run:
- `bun test stats/src/lib/dashboard-data.test.ts`
- `bun run typecheck`
- `bun run test:fast`
- `bun run build`
- `bun run test:env`
- `bun run test:smoke:dist`
<!-- SECTION:FINAL_SUMMARY:END -->
