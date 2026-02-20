---
id: TASK-92
title: Store Jellyfin token in encrypted local token store like AniList
status: Done
assignee: []
created_date: '2026-02-20 02:22'
updated_date: '2026-02-20 02:22'
labels: []
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Persist Jellyfin access token in local encrypted token storage and resolve `jellyfin.accessToken` from config-or-store, matching AniList token handling pattern.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Jellyfin login/setup saves token to encrypted local store instead of persisting in config.
- [ ] #2 Jellyfin logout clears local stored token.
- [ ] #3 Runtime resolves Jellyfin token from config override first, then stored token fallback when config value is blank.
- [ ] #4 Jellyfin runtime tests cover new token-store save/clear/fallback behavior.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added `src/core/services/jellyfin-token-store.ts` using Electron `safeStorage` with encrypted payload + plaintext fallback/migration behavior matching AniList store pattern. Wired token store in `src/main.ts` and changed Jellyfin flows so login/setup save token in store and write blank `jellyfin.accessToken` to config, while logout clears store. Updated resolved Jellyfin config path (`src/main/runtime/jellyfin-client-info.ts`) to apply config-first + stored-token fallback. Updated runtime deps/tests: `jellyfin-client-info*`, `jellyfin-cli-auth*`, `jellyfin-cli-main-deps*`, `jellyfin-setup-window*`, and `jellyfin-setup-window-main-deps*`. Updated docs/config notes to reflect local encrypted token storage behavior.
<!-- SECTION:NOTES:END -->
