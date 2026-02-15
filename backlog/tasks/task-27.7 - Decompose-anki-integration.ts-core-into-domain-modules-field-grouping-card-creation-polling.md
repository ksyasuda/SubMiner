---
id: TASK-27.7
title: >-
  Decompose anki-integration.ts core into domain modules (field-grouping,
  card-creation, polling)
status: To Do
assignee: []
created_date: '2026-02-15 07:00'
labels:
  - refactor
  - anki
  - architecture
dependencies:
  - TASK-27.3
references:
  - src/anki-integration.ts
  - src/anki-integration-duplicate.ts
  - src/anki-integration-ui-feedback.ts
  - src/anki-integration/ai.ts
parent_task_id: TASK-27
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
TASK-27.3 extracted leaf clusters (duplicate-detection 102 LOC, ai-translation 158 LOC, ui-feedback 107 LOC) but the core class remains at 2935 LOC. The heavy decomposition from the original TASK-27.3 plan was never executed.

Remaining extractions from the original plan:
1. **field-grouping** (~900 LOC) — `triggerFieldGroupingForLastAddedCard`, `applyFieldGrouping`, `computeFieldGroupingMergedFields`, `buildFieldGroupingPreview`, `performFieldGroupingMerge`, `handleFieldGroupingAuto`, `handleFieldGroupingManual`, plus ~15 span/parse/normalize helpers
2. **card-creation** (~350 LOC) — `createSentenceCard`, `setCardTypeFields`, `extractFields`, `processSentence`, field resolution helpers
3. **polling/lifecycle** (~250 LOC) — `start`, `stop`, `poll`, `pollOnce`, `processNewCard`

Also consolidate the scattered extraction files into `src/anki-integration/`:
- `src/anki-integration-duplicate.ts` → `src/anki-integration/duplicate.ts`
- `src/anki-integration-ui-feedback.ts` → `src/anki-integration/ui-feedback.ts`
- `src/anki-integration/ai.ts` (already there)
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 anki-integration.ts reduced below 800 LOC (facade + private state wiring only)
- [ ] #2 Field-grouping cluster (~900 LOC) extracted as its own module under src/anki-integration/
- [ ] #3 Card-creation and polling/lifecycle extracted as separate modules
- [ ] #4 All extracted modules consolidated under src/anki-integration/ directory
- [ ] #5 Existing facade API preserved — external callers unchanged
- [ ] #6 All existing tests pass; build compiles cleanly
<!-- AC:END -->
