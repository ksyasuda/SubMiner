---
id: TASK-93
title: Remove Jellyfin token/userId from config; use env override and stored session
status: Done
assignee: []
created_date: '2026-02-21 04:20'
updated_date: '2026-02-21 04:27'
labels: []
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Remove `jellyfin.accessToken` and `jellyfin.userId` from editable config surface. Resolve Jellyfin auth via env override (`SUBMINER_JELLYFIN_ACCESS_TOKEN`, optional `SUBMINER_JELLYFIN_USER_ID`) and locally stored Jellyfin auth session payload persisted by login/setup.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 `jellyfin.accessToken` and `jellyfin.userId` no longer parsed/documented as config fields.
- [ ] #2 Jellyfin login/setup persists token + userId in local encrypted session store.
- [ ] #3 Jellyfin runtime resolves auth from env override first, then stored session payload.
- [ ] #4 Jellyfin logout clears stored session payload.
- [ ] #5 Runtime tests cover env override + stored session fallback behavior.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented: evolved `jellyfin-token-store` to encrypted session payload store (`accessToken` + `userId`), updated Jellyfin resolver precedence to env-first (`SUBMINER_JELLYFIN_ACCESS_TOKEN`, optional `SUBMINER_JELLYFIN_USER_ID`) then stored session fallback, removed `jellyfin.accessToken`/`jellyfin.userId` from config defaults+parsing+public docs/example, and updated CLI/setup auth wiring to persist/clear session store while keeping `serverUrl`/`username` config behavior.
<!-- SECTION:NOTES:END -->
