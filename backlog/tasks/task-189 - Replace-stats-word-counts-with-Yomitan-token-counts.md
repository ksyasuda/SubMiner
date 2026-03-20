---
id: TASK-189
title: Replace stats word counts with Yomitan token counts
status: Done
assignee:
  - codex
created_date: '2026-03-18 01:35'
updated_date: '2026-03-18 05:28'
labels:
  - stats
  - tokenizer
  - bug
milestone: m-1
dependencies: []
references:
  - src/core/services/immersion-tracker-service.ts
  - src/core/services/immersion-tracker/reducer.ts
  - src/core/services/immersion-tracker/storage.ts
  - src/core/services/immersion-tracker/query.ts
  - src/core/services/immersion-tracker/lifetime.ts
  - stats/src/components
  - stats/src/lib/yomitan-lookup.ts
priority: medium
ordinal: 100500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Replace heuristic immersion stats word counting with Yomitan token counts. Session/media/anime stats should use the exact merged Yomitan token stream as the denominator and display metric, with no whitespace/CJK-character fallback and no active `wordsSeen` concept in the runtime, storage, API, or stats UI.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `recordSubtitleLine` derives session count deltas from Yomitan token arrays instead of `calculateTextMetrics`.
- [x] #2 Active immersion tracking/storage/query code no longer depends on `wordsSeen` / `totalWordsSeen` fields for stats behavior.
- [x] #3 Stats UI labels and lookup-rate copy refer to tokens instead of words where those counts are shown to users.
- [x] #4 Regression tests cover token-count sourcing, zero-count behavior when tokenization payload is absent, and updated stats copy.
- [x] #5 A changelog fragment documents the user-visible stats denominator change.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing tracker tests proving subtitle count metrics come from Yomitan token arrays and stay zero when tokenization is absent.
2. Add failing stats UI tests for token-based copy and token-count display helpers.
3. Remove `wordsSeen` from active tracker/session/query/type paths and use `tokensSeen` as the single stats count field.
4. Update stats UI labels and lookup-rate copy from words to tokens.
5. Run targeted verification, then add the changelog fragment and any needed docs update.
<!-- SECTION:PLAN:END -->

## Outcome

<!-- SECTION:OUTCOME:BEGIN -->
Completed. Stats subtitle counts now come directly from Yomitan merged-token counts, `wordsSeen` is removed from the active tracker/storage/query/UI path, token-facing copy is updated, and focused regression coverage plus `bun run typecheck` are green.
<!-- SECTION:OUTCOME:END -->
