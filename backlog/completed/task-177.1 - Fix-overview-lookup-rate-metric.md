---
id: TASK-177.1
title: Fix overview lookup rate metric
status: Done
assignee:
  - '@codex'
created_date: '2026-03-19 17:46'
updated_date: '2026-03-23 03:22'
labels:
  - stats
  - immersion-tracking
  - yomitan
dependencies: []
references:
  - stats/src/components/overview/OverviewTab.tsx
  - stats/src/lib/dashboard-data.ts
  - stats/src/lib/yomitan-lookup.ts
  - src/core/services/immersion-tracker/query.ts
  - src/core/services/stats-server.ts
parent_task_id: TASK-177
priority: medium
ordinal: 132500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Update the stats homepage Tracking Snapshot so Lookup Rate reflects lifetime intentional Yomitan lookups normalized by total tokens seen, matching the newer stats semantics already used in session, media, and anime views.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Overview data exposes the lifetime totals needed to compute global Yomitan lookups per 100 tokens on the homepage
- [x] #2 The homepage Tracking Snapshot Lookup Rate card shows Yomitan lookup rate as `X / 100 tokens` with tooltip/copy aligned to that meaning
- [x] #3 Automated tests cover the lifetime totals plumbing and homepage summary/rendering change
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Extend overview lifetime hints/query plumbing to include total tokens seen and total intentional Yomitan lookups from finished sessions.
2. Add/adjust focused tests first for query hints, stats overview API typing/mocks, and overview summary formatting so the homepage metric fails under old semantics.
3. Update the overview summary/card to derive Lookup Rate from lifetime Yomitan lookups per 100 tokens and align tooltip/copy with that meaning.
4. Run focused verification on the touched query, stats-server, and stats UI tests; record results and blockers in the task notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Extended overview lifetime hints to include total tokens seen and total intentional Yomitan lookups from finished sessions so the homepage can compute a true global lookup rate.

Extracted the homepage Tracking Snapshot into a dedicated presentational component to keep OverviewTab smaller and make the Lookup Rate card copy directly testable.

Focused verification passed for query hints, IPC/stats overview plumbing, stats server overview response, dashboard summary logic, and homepage snapshot rendering.

SubMiner verifier core lane artifact: .tmp/skill-verification/subminer-verify-20260319-105320-7FDlwh. `bun run typecheck` passed there; `bun run test:fast` failed for a pre-existing/unrelated environment issue in scripts/update-aur-package.test.ts because scripts/update-aur-package.sh reported `mapfile: command not found`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Homepage Lookup Rate now uses lifetime intentional Yomitan lookups normalized by lifetime tokens seen, matching the existing session/media/anime semantics instead of the old known-word hit-rate metric. I extended overview query hints and API typings with total token and Yomitan lookup totals, updated the overview summary builder to reuse the shared per-100-token formatter, and replaced the inline Tracking Snapshot block with a dedicated component that renders `X / 100 tokens` plus Yomitan-specific tooltip copy.

Tests added/updated: query hints coverage for the new lifetime totals, stats server and IPC overview fixtures, overview summary assertions, and a dedicated Tracking Snapshot render test for the homepage card text. Focused `bun test` runs passed for those touched areas. Repo-native verifier `--lane core` also passed `bun run typecheck`; its `bun run test:fast` step still fails for the unrelated existing `scripts/update-aur-package.sh: line 71: mapfile: command not found` environment issue.
<!-- SECTION:FINAL_SUMMARY:END -->
