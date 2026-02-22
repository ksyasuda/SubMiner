---
id: TASK-29.3
title: Fix AniList OAuth callback token handling in setup window
status: Done
assignee: []
created_date: '2026-02-19 16:56'
updated_date: '2026-02-22 07:49'
labels:
  - anilist
  - oauth
  - bug
dependencies: []
parent_task_id: TASK-29
priority: high
ordinal: 97000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
AniList login flow shows unsupported_grant_type after auth because setup window does not consume callback URL token and persist it.

Need robust callback handling for both query and hash access_token forms and graceful close/success UX.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 AniList setup flow persists access token from callback URL query/hash
- [x] #2 Setup window closes and state updates to resolved when token captured
- [x] #3 No unsafe navigation regressions in AniList setup window
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
<!-- SECTION:IMPLEMENTATION_NOTES:BEGIN -->
- Added robust AniList callback token parsing for query/hash and `subminer://anilist-setup?...` deep links.
- Added app-level protocol handling (`open-url` + second-instance argv deep link parsing) so browser callback buttons resolve even when setup window is not navigating.
- Added/updated targeted AniList setup runtime tests and verified build + runtime test pass.
<!-- SECTION:IMPLEMENTATION_NOTES:END -->
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Build passes
- [x] #2 Targeted AniList runtime tests pass
<!-- DOD:END -->
