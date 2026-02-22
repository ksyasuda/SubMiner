---
id: TASK-102
title: Restore Jellyfin session migration contract and build green
status: Done
assignee: []
created_date: '2026-02-22 07:12'
updated_date: '2026-02-22 07:51'
labels:
  - bug
  - jellyfin
  - maintainability
dependencies:
  - TASK-93
priority: high
ordinal: 99000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`bun run build` is currently failing after refactor/rebase drift in Jellyfin auth/session wiring.

Observed drift:
- `src/config/definitions/defaults-integrations.ts` still includes `jellyfin.accessToken` and `jellyfin.userId`.
- `src/config/resolve/integrations.ts` still parses those removed config keys.
- `src/main.ts` still calls removed token-store methods (`loadToken/saveToken/clearToken`) instead of session-store methods.
- setup/config patch flows still attempt to write removed Jellyfin auth fields to config.

This task restores the TASK-93 contract end-to-end: auth in env/stored session, no token/userId in editable config.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Remove stale `jellyfin.accessToken`/`jellyfin.userId` keys from defaults and config resolve parsing.
2. Update `src/main.ts` Jellyfin deps wiring to use `loadStoredSession/saveStoredSession/clearStoredSession` with `createJellyfinTokenStore`.
3. Update Jellyfin setup/auth flows so config patch writes only supported Jellyfin config keys.
4. Update launcher-side Jellyfin config loader to stop reading removed auth keys from config file.
5. Add/adjust regression tests around Jellyfin resolver precedence and setup/auth save+clear paths.
6. Verify with `bun run build`, `bun run test:config:src`, and targeted Jellyfin runtime tests.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 `bun run build` passes on `main` after Jellyfin session contract fix.
- [ ] #2 No production config defaults/resolvers parse `jellyfin.accessToken` or `jellyfin.userId`.
- [ ] #3 Main/runtime Jellyfin auth wiring compiles against session-store API only.
- [ ] #4 Jellyfin setup/login/logout paths are covered by source tests and pass.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Compile + config test lanes pass without local patches.
- [ ] #2 Runtime Jellyfin tests include session-store save/load/clear expectations.
- [ ] #3 Docs/examples reflect session-store contract and no longer mention removed config auth keys.
<!-- DOD:END -->
