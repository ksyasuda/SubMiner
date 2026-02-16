---
id: TASK-53
title: Consolidate fragmented JLPT utility modules
status: Done
assignee: []
created_date: '2026-02-16 04:47'
updated_date: '2026-02-16 04:57'
labels: []
dependencies: []
references:
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/core/services/jlpt-token-filter-config.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/core/services/jlpt-excluded-terms.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/core/services/jlpt-ignored-mecab-pos1.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
The JLPT-related functionality is currently split across three tiny files that create unnecessary navigation overhead and add cognitive load when working with JLPT token filtering.

Files to consolidate:
- `src/core/services/jlpt-token-filter-config.ts` (24 lines)
- `src/core/services/jlpt-excluded-terms.ts` (30 lines)  
- `src/core/services/jlpt-ignored-mecab-pos1.ts` (46 lines)

These three files are tightly coupled (one imports from another) and all deal with JLPT token filtering logic. They should be merged into a single `jlpt-token-filter.ts` module with clear section comments.

Benefits:
- Reduces file count by 2
- Eliminates circular import potential
- Easier to understand JLPT filtering logic holistically
- Simplifies barrel exports from services/index.ts
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Merge jlpt-token-filter-config.ts, jlpt-excluded-terms.ts, and jlpt-ignored-mecab-pos1.ts into a single jlpt-token-filter.ts file
- [ ] #2 Update all imports across the codebase to use the consolidated module
- [ ] #3 Update services/index.ts barrel exports
- [ ] #4 Ensure all existing tests pass
- [ ] #5 Verify no functionality is lost in consolidation
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Consolidated JLPT filtering utilities into `src/core/services/jlpt-token-filter.ts` and updated all imports to use the new module. Added consolidated exports to `src/core/services/index.ts` for the JLPT token filter API (`shouldIgnoreJlptForMecabPos1`, `shouldIgnoreJlptByTerm`, POS metadata, etc.). Deleted the three fragmented files:
- `jlpt-token-filter-config.ts`
- `jlpt-excluded-terms.ts`
- `jlpt-ignored-mecab-pos1.ts`

Validation:
- `pnpm exec tsc --noEmit` passes after the refactor.
- `node --test dist/core/services/tokenizer-service.test.js` runs with 19/20 passing; one failure (`tokenizeSubtitleService uses Yomitan parser result when available`) appears to be pre-existing in current tree and matches earlier full test run failure.
- Full `pnpm run test:fast` still fails, but failures are outside this consolidation scope (e.g., overlay-shortcut-handler logger message shape and tokenizer Yomitan-parse-path test), consistent with baseline failures in this working state.
<!-- SECTION:FINAL_SUMMARY:END -->
