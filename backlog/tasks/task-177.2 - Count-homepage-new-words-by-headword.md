---
id: TASK-177.2
title: Count homepage new words by headword
status: Done
assignee:
  - '@codex'
created_date: '2026-03-19 19:38'
updated_date: '2026-03-19 19:40'
labels:
  - stats
  - immersion-tracking
  - vocabulary
dependencies: []
references:
  - src/core/services/immersion-tracker/query.ts
  - stats/src/components/overview/TrackingSnapshot.tsx
  - stats/src/lib/dashboard-data.ts
parent_task_id: TASK-177
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Align the homepage New Words metric with the Known Words semantics by counting distinct headwords first seen in the selected window, so inflected or alternate forms of the same word do not inflate the summary.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Homepage new-word counts use distinct headwords by earliest first-seen timestamp instead of counting separate word-form rows
- [x] #2 Homepage tooltip/copy reflects the headword-based semantics
- [x] #3 Automated tests cover the headword de-duplication behavior and affected overview copy
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Change the new-word aggregate query to group `imm_words` by headword, compute each headword's earliest `first_seen`, and count headwords whose first sighting falls within today/week windows.
2. Add failing tests first for the aggregate path so multiple rows sharing a headword only contribute once.
3. Update homepage tooltip/copy to say unique headwords first seen today/week.
4. Run focused query and stats overview tests, then record verification and any blockers.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Updated the new-word aggregate to count distinct headwords by each headword's earliest `first_seen` timestamp, so multiple inflected/form rows for the same headword contribute only once.

Adjusted homepage tooltip copy to say unique headwords first seen today/week, keeping the visible card labels unchanged.

Focused verification passed for the query aggregate and homepage snapshot tests.

SubMiner verifier core lane artifact: .tmp/skill-verification/subminer-verify-20260319-123942-4intgW. `bun run typecheck` passed there; `bun run test:fast` still fails for the unrelated environment issue in scripts/update-aur-package.test.ts (`mapfile: command not found`).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Homepage New Words now uses headword-level semantics instead of counting separate `(headword, word, reading)` rows. The aggregate query groups `imm_words` by headword, uses each headword's earliest `first_seen`, and counts headwords first seen today or this week so alternate forms do not inflate the summary. The homepage tooltip copy now explicitly says the metric is based on unique headwords.

Added focused regression coverage for the de-duplication rule in `getQueryHints` and for the updated homepage tooltip text. Targeted `bun test` runs passed for the touched query and stats UI files. Repo verifier `--lane core` again passed `bun run typecheck`; `bun run test:fast` remains blocked by the unrelated existing `scripts/update-aur-package.sh: line 71: mapfile: command not found` failure.
<!-- SECTION:FINAL_SUMMARY:END -->
