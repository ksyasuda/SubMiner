---
id: TASK-27.3
title: Refactor anki-integration.ts into domain-specific service modules
status: To Do
assignee:
  - backend
created_date: '2026-02-13 17:13'
updated_date: '2026-02-13 21:13'
labels:
  - 'owner:backend'
dependencies:
  - TASK-27.1
references:
  - src/anki-integration.ts
documentation:
  - docs/architecture.md
parent_task_id: TASK-27
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Split anki-integration.ts (2,679 LOC, 60+ methods) into cohesive modules with clear handoff contracts. Keep a stable facade API to avoid broad call-site churn.

## Target Decomposition (7 modules)

Based on method clustering analysis:

1. **polling/lifecycle** (~250 LOC) — `start`, `stop`, `poll`, `pollOnce`, `processNewCard`
2. **card-creation** (~350 LOC) — `createSentenceCard`, `setCardTypeFields`, `extractFields`, `processSentence`, field resolution helpers
3. **media-generation** (~200 LOC) — `generateAudio`, `generateImage`, filename generation, `generateMediaForMerge`
4. **field-grouping** (~900 LOC) — `triggerFieldGroupingForLastAddedCard`, `applyFieldGrouping`, `computeFieldGroupingMergedFields`, `buildFieldGroupingPreview`, `performFieldGroupingMerge`, `handleFieldGroupingAuto`, `handleFieldGroupingManual`, plus ~15 span/parse/normalize helpers
5. **duplicate-detection** (~100 LOC) — `findDuplicateNote`, `findFirstExactDuplicateNoteId`, search escaping
6. **ai-translation** (~100 LOC) — `translateSentenceWithAi`, `extractAiText`, `normalizeOpenAiBaseUrl`
7. **ui-feedback** (~150 LOC) — progress tracking, OSD notifications, status notifications

Plus a **facade** (`src/anki-integration/index.ts`) that re-exports the public API for backward compatibility.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Extract at least 6 modules matching the decomposition map above (7 recommended).
- [ ] #2 Keep a stable facade API in src/anki-integration/index.ts so external callers (main.ts, mining-service) don't change.
- [ ] #3 Field grouping cluster (~900 LOC) is extracted as its own module — it's the largest single concern.
- [ ] #4 AI/translation integration is extracted separately from card creation.
- [ ] #5 Each module defines explicit input/output types; no module exceeds 400 LOC unless justified in a TODO comment.
- [ ] #6 The 15 private state fields (pollingInterval, backoffMs, progressTimer, etc.) are managed through a shared internal state object or passed explicitly — not scattered across files as module-level lets.
- [ ] #7 Preserve existing external behavior for all config keys and session flow.
- [ ] #8 All existing tests pass; add focused unit tests for at least the field-grouping and card-creation modules.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
## Execution Notes

This task is self-contained — anki-integration.ts is a single class with a clear public API consumed by main.ts and mining-service. Internal restructuring doesn't affect other files as long as the facade is maintained.

**Should run first in Phase 2** because:
- It's the largest file (2,679 LOC vs 1,392 for main.ts)
- Its public API is narrow (class constructor + ~5 public methods)
- main.ts instantiates AnkiIntegration, so stabilizing its API before splitting main.ts avoids double-refactoring

## Key Risk
The class has 15 private state fields that create implicit coupling between methods. The `updateLastAddedFromClipboard` method alone is ~230 lines and touches polling state, media generation, and card updates. Extraction order matters: start with the leaf clusters (ai-translation, ui-feedback, duplicate-detection) and work inward toward the stateful core (polling, card-creation, field-grouping).
<!-- SECTION:NOTES:END -->
