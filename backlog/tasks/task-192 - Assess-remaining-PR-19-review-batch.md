---
id: TASK-192
title: 'Assess remaining PR #19 review batch'
status: Done
assignee:
  - codex
created_date: '2026-03-17 23:24'
updated_date: '2026-03-17 23:42'
labels:
  - pr-review
  - stats
  - docs
milestone: m-1
dependencies: []
references:
  - docs/superpowers/plans/2026-03-12-immersion-stats-page.md
  - src/core/services/immersion-tracker/__tests__/query.test.ts
  - src/core/services/ipc.ts
  - src/core/services/stats-server.ts
  - src/main.ts
  - src/renderer/handlers/keyboard.ts
  - stats/src
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Validate the remaining PR #19 automated review findings against the current branch, implement only the technically correct fixes, and document which comments are stale, already addressed, or not warranted.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Each remaining review comment is classified as actionable, already fixed, stale, or not warranted
- [x] #2 Confirmed bugs or correctness issues are fixed with focused regression coverage where it fits
- [x] #3 Final notes record which comments were intentionally not applied and why
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Inspect the referenced files in batches and compare each comment against current branch behavior.
2. Separate correctness/security regressions from stylistic nitpicks and already-fixed items.
3. Add tests first for confirmed behavior bugs where practical, apply the smallest safe fixes, and rerun targeted verification.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Swept the pasted PR #19 review batch against the current branch.

Classification:
- Already fixed on current branch: `src/core/services/immersion-tracker/__tests__/query.test.ts` cleanup rethrow, `src/core/services/ipc.ts` limit validation, `src/core/services/stats-server.ts` max-limit parsing and CORS removal, `src/main.ts` quit-path TDZ issue, `src/renderer/handlers/keyboard.ts` stats-toggle shortcut ordering/config usage, `stats/src/components/vocabulary/WordList.tsx`, `stats/src/hooks/useSessions.ts`, `stats/src/hooks/useTrends.ts` stale-error reset, `src/core/services/__tests__/stats-server.test.ts` kanji endpoint/readability notes, `src/core/services/stats-window.ts`, `stats/src/App.tsx`, `stats/src/components/layout/TabBar.tsx`, `stats/src/components/overview/QuickStats.tsx`, `stats/src/components/overview/WatchTimeChart.tsx`, `stats/src/components/sessions/SessionDetail.tsx`, `stats/src/components/sessions/SessionRow.tsx`, `stats/src/components/trends/DateRangeSelector.tsx`, `stats/src/components/vocabulary/KanjiBreakdown.tsx`, `stats/src/components/vocabulary/VocabularyTab.tsx`, `stats/src/hooks/useVocabulary.ts`, `stats/src/lib/api-client.ts`, `stats/src/types/stats.ts`.
- Stale / obsolete against current architecture: `docs/superpowers/plans/2026-03-12-immersion-stats-page.md` path does not exist on this branch; `stats/src/components/trends/TrendsTab.tsx` / monthly-range comments describe older client-side aggregation code that is no longer present because trends now come from `getTrendsDashboard`.
- Not warranted as written: `stats/src/lib/formatters.ts` no longer emits negative `Xd ago`; current code short-circuits future timestamps to `just now`, so the reported bug condition is gone even though the suggested wording differs.
- Actionable and fixed now: `src/core/services/ipc.ts` no-tracker `statsGetOverview` fallback omitted required hint fields (`totalLookupCount`, `totalLookupHits`, `newWordsToday`, `newWordsThisWeek`). Added the missing fields in the fallback object and updated IPC tests to assert the full shape.

Verification:
- `bun test src/core/services/ipc.test.ts`
- `bun test src/core/services/ipc.test.ts --test-name-pattern "empty stats overview shape without a tracker|validates and clamps stats request limits"`
- `bash .agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh src/core/services/ipc.ts src/core/services/ipc.test.ts`

Repo verifier note:
- `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane core src/core/services/ipc.ts src/core/services/ipc.test.ts`
- That verifier run captured a temporary `bun run typecheck` failure in `src/anki-integration.test.ts` and `src/core/services/__tests__/stats-server.test.ts`, but a fresh rerun after the follow-up validation no longer reproduces those diagnostics.
- Fresh verification: `bun run typecheck` passes locally.
- artifact dir from the earlier failed verifier snapshot: `.tmp/skill-verification/subminer-verify-20260317-234027-i6QJ3n`
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
The larger pasted PR #19 review batch was not mostly new work on the current branch. After verifying each item against the live code, almost all were already fixed or stale. One additional item was still actionable: the no-tracker fallback returned by `statsGetOverview` in `src/core/services/ipc.ts` omitted required hint fields, which made the fallback shape inconsistent with the normal overview payload. That fallback is now fixed and covered by IPC tests.

Count-wise: the earlier open CodeRabbit service comments contributed 2 actionable fixes, and this larger pasted batch contributed 1 additional actionable fix on top of those.
<!-- SECTION:FINAL_SUMMARY:END -->
