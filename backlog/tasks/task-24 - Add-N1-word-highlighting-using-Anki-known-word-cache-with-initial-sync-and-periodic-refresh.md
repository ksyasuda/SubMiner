---
id: TASK-24
title: >-
  Add N+1 word highlighting using Anki-known-word cache with initial sync and
  periodic refresh
status: To Do
assignee: []
created_date: '2026-02-13 16:45'
labels: []
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Implement subtitle highlighting for words already known in Anki (N+1 workflow support) by introducing a one-time bootstrap query of the user’s Anki known-word set, storing it locally, and refreshing it periodically to reflect deck updates. The feature should allow fast in-session lookups to determine known words and visually distinguish them in subtitle rendering.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Add an opt-in setting/feature flag for N+1 highlighting and default it to disabled for backward-compatible behavior.
- [ ] #2 Implement a one-time import/sync that queries known-word data from Anki into a local store on first enable or explicit refresh.
- [ ] #3 Store known words locally in an efficient structure for fast lookup during subtitle rendering.
- [ ] #4 Run periodic refresh on a configurable interval and expose a manual refresh action.
- [ ] #5 Ensure local cache updates replace or merge safely without corrupting in-flight subtitle rendering queries.
- [ ] #6 Known/unknown lookup decisions are applied consistently to subtitle tokens for highlighting without impacting tokenization performance.
- [ ] #7 Non-targeted words remain visually unchanged and all existing subtitle interactions remain unaffected.
- [ ] #8 Add tests/validation for initial sync success, refresh update, and disabled-mode no-lookup behavior.
- [ ] #9 Document Anki data source expectations, failure handling, and update policy/interval behavior.
- [ ] #10 If full Anki query integration is not possible in this environment, define deterministic fallback behavior with clear user-visible messaging.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 N+1 known-word highlighting is configurable, performs local cached lookups, and is demonstrated to update correctly after periodic/manual refresh.
<!-- DOD:END -->
