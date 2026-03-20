---
id: TASK-183
title: Fix blank stats vocabulary page regression
status: Done
milestone: m-1
assignee:
  - codex
created_date: '2026-03-17 16:23'
updated_date: '2026-03-18 05:28'
labels: []
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/stats/src/components/vocabulary/VocabularyTab.tsx
  - /Users/sudacode/projects/japanese/SubMiner/stats/src/App.tsx
  - /Users/sudacode/projects/japanese/SubMiner/stats/src/lib/api-client.ts
ordinal: 108500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Diagnose and fix the stats dashboard regression where the Vocabulary tab renders blank at runtime. Capture the frontend failure with browser debugging, add regression coverage, and restore the vocabulary view.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Vocabulary tab renders without a blank-screen failure in the stats dashboard
- [x] #2 Automated test coverage reproduces the failing code path and passes with the fix
- [x] #3 Targeted verification covers the affected stats UI/runtime path
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Reproduce the blank Vocabulary tab locally with a browser-visible stats UI instance and capture console/network failure details.
2. Add a focused regression test for the failing Vocabulary tab code path before editing production code.
3. Implement the minimal fix in the stats UI/runtime path.
4. Re-run targeted browser and automated verification, then record any skipped broader checks.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Identified the runtime failure in the browser console: React reported a hook-order change in `VocabularyTab` after the tab moved from loading to loaded state (`Rendered more hooks than during the previous render`).

Fixed `stats/src/components/vocabulary/VocabularyTab.tsx` by removing the late `useMemo` hook and computing `knownWordCount` as a plain derived value after the loading/error guards.

Added regression coverage in `stats/src/lib/vocabulary-tab.test.ts` to assert that `VocabularyTab` declares all hooks before the loading/error early returns.

Verification: `bun test stats/src/lib/vocabulary-tab.test.ts`, `bun test stats/src/lib/yomitan-lookup.test.tsx`, `bun run build:stats`, and a live Playwright check against the Vite app with stubbed stats API data all passed.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed the blank Vocabulary tab regression in the stats UI. The root cause was a late `useMemo` hook declared after the loading/error early returns in `VocabularyTab`, which caused React to crash once vocabulary data finished loading. Removed that late hook, added a regression test guarding hook placement, verified the stats bundle builds, and confirmed in a live browser that the Vocabulary tab now renders loaded content instead of white-screening.
<!-- SECTION:FINAL_SUMMARY:END -->
