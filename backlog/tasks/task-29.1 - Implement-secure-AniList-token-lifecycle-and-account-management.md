---
id: TASK-29.1
title: Implement secure AniList token lifecycle and account management
status: Done
assignee: []
created_date: '2026-02-17 04:12'
updated_date: '2026-02-17 04:19'
labels:
  - anilist
  - security
  - auth
dependencies: []
parent_task_id: TASK-29
priority: medium
---

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Access token is stored in secure local storage rather than plain config.
- [ ] #2 Token connect/disconnect UX supports revocation/logout and re-auth setup.
- [ ] #3 Startup flow validates token presence/state and surfaces actionable errors.
- [ ] #4 Docs describe token management and security expectations.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented secure AniList token lifecycle: config token persists to encrypted token store, stored token fallback is auto-resolved at runtime, and auth state source now distinguishes literal vs stored.
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Token lifecycle module wired into AniList update/auth flow.
- [ ] #2 Unit/integration coverage added for token storage and logout paths.
<!-- DOD:END -->
