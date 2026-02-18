---
id: TASK-24
title: >-
  Add N+1 word highlighting using Anki-known-word cache with initial sync and
  periodic refresh
status: Done
assignee: []
created_date: '2026-02-13 16:45'
updated_date: '2026-02-18 04:11'
labels: []
dependencies: []
priority: high
ordinal: 37000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Implement subtitle highlighting for words already known in Anki (N+1 workflow support) by introducing a one-time bootstrap query of the user’s Anki known-word set, storing it locally, and refreshing it periodically to reflect deck updates. The feature should allow fast in-session lookups to determine known words and visually distinguish them in subtitle rendering.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Add an opt-in setting/feature flag for N+1 highlighting and default it to disabled for backward-compatible behavior.
- [x] #2 Implement a one-time import/sync that queries known-word data from Anki into a local store on first enable or explicit refresh.
- [x] #3 Store known words locally in an efficient structure for fast lookup during subtitle rendering.
- [x] #4 Run periodic refresh on a configurable interval and expose a manual refresh action.
- [x] #5 Ensure local cache updates replace or merge safely without corrupting in-flight subtitle rendering queries.
- [x] #6 Known/unknown lookup decisions are applied consistently to subtitle tokens for highlighting without impacting tokenization performance.
- [x] #7 Non-targeted words remain visually unchanged and all existing subtitle interactions remain unaffected.
- [x] #8 Add tests/validation for initial sync success, refresh update, and disabled-mode no-lookup behavior.
- [x] #9 Document Anki data source expectations, failure handling, and update policy/interval behavior.
- [x] #10 If full Anki query integration is not possible in this environment, define deterministic fallback behavior with clear user-visible messaging.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented in refactor via merge from task-24-known-word-refresh (commits 854b8fb, e8f2431, ed5a249). Includes manual/periodic known-word cache refresh, opt-in N+1 highlighting path, cache persistence behavior, CLI refresh command, and related tests/docs updates.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 N+1 known-word highlighting is configurable, performs local cached lookups, and is demonstrated to update correctly after periodic/manual refresh.
<!-- DOD:END -->
