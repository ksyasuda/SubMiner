---
id: TASK-355
title: Unify AniList API throttling across dictionary stats and tracking
status: In Progress
assignee: []
created_date: '2026-05-12 21:49'
updated_date: '2026-05-13 01:21'
labels:
  - anilist
  - rate-limit
  - bug
dependencies: []
references:
  - 'https://docs.anilist.co/guide/rate-limiting'
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Audit and fix AniList GraphQL usage so character dictionary generation, stats search/cover art, and post-watch tracking share conservative request pacing and honor AniList rate-limit response headers. Current logs do not show 429s, but source has separate/unthrottled call paths and repeated dictionary lookup failures for the same title.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 All AniList GraphQL call paths use a shared/conservative limiter or equivalent pacing before requests.
- [ ] #2 429 responses honor Retry-After/X-RateLimit-Reset and do not continue hammering the API.
- [x] #3 Stats AniList search endpoint no longer bypasses the AniList rate limiter.
- [x] #4 Post-watch tracking no longer bypasses the AniList rate limiter.
- [x] #5 Focused regression tests cover limiter use for stats search and post-watch tracking, plus existing limiter behavior remains green.
- [x] #6 Stats cover-art lookup reuses already stored AniList cover data for the same anime before issuing another AniList API request.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented stats cover-art cache reuse across videos in the same anime before any AniList/image fetch. Added limiter plumbing for stats manual AniList search and post-watch tracking; both paths now call acquire before GraphQL and record response headers afterward. Character dictionary still uses its existing local pacing and remains follow-up work for fully shared limiter/header handling.
<!-- SECTION:NOTES:END -->
