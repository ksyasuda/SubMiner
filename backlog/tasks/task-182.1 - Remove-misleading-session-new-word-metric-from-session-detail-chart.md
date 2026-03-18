---
id: TASK-182.1
title: Remove misleading session new-word metric from session detail chart
status: Done
assignee:
  - '@codex'
created_date: '2026-03-18 01:41'
updated_date: '2026-03-18 01:45'
labels:
  - bug
  - stats
  - ui
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/stats/src/components/sessions/SessionDetail.tsx
  - >-
    /Users/sudacode/projects/japanese/SubMiner/stats/src/components/sessions/SessionsTab.tsx
  - >-
    /Users/sudacode/projects/japanese/SubMiner/stats/src/lib/media-session-list.test.tsx
parent_task_id: TASK-182
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Remove the misleading `New words` series from the session detail chart so the stats UI no longer presents a fabricated metric that mirrors total words. Keep the session chart focused on the real cumulative totals already backed by tracker data.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Expanded session detail chart no longer renders or labels a `New words` metric in the graph, tooltip, or legend
- [x] #2 Session detail still renders total-word and known-word series correctly after the metric removal
- [x] #3 Automated frontend coverage prevents the `New words` label from reappearing in expanded session detail
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a focused stats frontend regression test that renders expanded session detail and asserts the misleading `New words` label is absent while `Total words` remains.
2. Remove the fabricated `New words` area series, tooltip mapping, legend chip, and now-unused left-axis chart plumbing from `stats/src/components/sessions/SessionDetail.tsx`.
3. Add a user-visible changelog fragment describing the session chart cleanup.
4. Run targeted frontend tests plus cheap verification and record any blockers.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added a focused server-render regression test for SessionDetail copy to ensure the misleading `New words` label stays removed.

Removed the fabricated `New words` chart series and its legend/tooltip plumbing from the expanded session detail view.

Verification: `bun test stats/src/lib/session-detail.test.tsx stats/src/lib/media-session-list.test.tsx` passed. `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane core stats/src/components/sessions/SessionDetail.tsx stats/src/lib/session-detail.test.tsx changes/2026-03-18-remove-session-new-words-series.md` passed and wrote artifacts under `.tmp/skill-verification/subminer-verify-20260317-184440-1aMWkM`.

Manual spot-check note: `bun test stats/src/lib/yomitan-lookup.test.tsx` is currently red on a pre-existing `AnimeOverviewStats` lookup-rate assertion unrelated to this session-detail change.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Removed the misleading `New words` metric from expanded session charts. Session detail now shows only the real total-word and known-word lines, backed by existing tracker data, with regression coverage that prevents the `New words` label from reappearing.
<!-- SECTION:FINAL_SUMMARY:END -->
