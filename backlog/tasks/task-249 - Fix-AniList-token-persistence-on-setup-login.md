---
id: TASK-249
title: Fix AniList token persistence on setup login
status: Done
assignee: []
created_date: '2026-03-29 10:08'
updated_date: '2026-03-31 19:37'
labels:
  - anilist
  - bug
dependencies: []
documentation:
  - src/main/runtime/anilist-setup.ts
  - src/core/services/anilist/anilist-token-store.ts
  - src/main/runtime/anilist-token-refresh.ts
  - docs-site/anilist-integration.md
priority: high
ordinal: 166500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
AniList setup can appear successful but the token is not persisted across restarts. Investigate the setup callback and token store path so the app either saves the token reliably or surfaces persistence failure instead of reopening setup on every launch.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 AniList setup login persists a usable token across app restarts when safeStorage works
- [ ] #2 If token persistence fails the setup flow reports the failure instead of pretending login succeeded
- [ ] #3 Regression coverage exists for the callback/save path and the refresh path that reopens setup when no token is available
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Pinned installed mpv plugin configs to the current SubMiner binary so standalone mpv launches reuse the same app identity that saved AniList tokens. Added startup self-heal for existing blank binary_path configs, install-time binary_path writes for fresh plugin installs, regression tests for both paths, and docs updates describing the new behavior.
<!-- SECTION:FINAL_SUMMARY:END -->
