# codex-jellyfin-secret-store-20260220T101428Z-om4z

- alias: `codex-jellyfin-secret-store`
- mission: `Move Jellyfin token/userId out of config into env override + stored session payload`
- status: `completed`
- last_update_utc: `2026-02-21T04:27:24Z`

## Intent

- compare AniList token persistence path vs Jellyfin token persistence path
- answer feasibility + current behavior from code

## Planned Files

- `src/main/runtime/anilist-token-refresh.ts`
- `src/main/runtime/*anilist*`
- `src/core/services/jellyfin.ts`
- `src/main/runtime/*jellyfin*`
- `src/config/*`
- docs refs if needed

## Assumptions

- user asking architecture/feasibility question; likely no code change requested yet

## Findings

- AniList token store present: `src/core/services/anilist/anilist-token-store.ts` (`safeStorage` encrypt/decrypt + persisted file)
- AniList runtime wiring present: `src/main.ts` creates `anilistTokenStore`; refresh/setup paths use `saveToken/loadToken`
- Jellyfin auth currently writes token directly into config via `patchRawConfig({ jellyfin: { accessToken } })`
- Docs confirm current behavior: Jellyfin token persisted in config (`docs/jellyfin-integration.md`)

## Outcome

- no code changes; answered feasibility/current state from repo
- implemented requested change: Jellyfin token now persisted in local encrypted token store with config override fallback

## Files Touched

- `src/core/services/jellyfin-token-store.ts`
- `src/main.ts`
- `src/main/runtime/jellyfin-client-info.ts`
- `src/main/runtime/jellyfin-client-info.test.ts`
- `src/main/runtime/jellyfin-client-info-main-deps.ts`
- `src/main/runtime/jellyfin-client-info-main-deps.test.ts`
- `src/main/runtime/jellyfin-cli-auth.ts`
- `src/main/runtime/jellyfin-cli-auth.test.ts`
- `src/main/runtime/jellyfin-cli-main-deps.ts`
- `src/main/runtime/jellyfin-cli-main-deps.test.ts`
- `src/main/runtime/jellyfin-setup-window.ts`
- `src/main/runtime/jellyfin-setup-window.test.ts`
- `src/main/runtime/jellyfin-setup-window-main-deps.ts`
- `src/main/runtime/jellyfin-setup-window-main-deps.test.ts`
- `docs/jellyfin-integration.md`
- `docs/configuration.md`
- `docs/public/config.example.jsonc`
- `src/config/definitions.ts`
- `backlog/tasks/task-92 - Store-Jellyfin-token-in-encrypted-local-token-store-like-AniList.md`

## Verification

- `bun run build`
- `node --test dist/main/runtime/jellyfin-client-info.test.js dist/main/runtime/jellyfin-cli-auth.test.js dist/main/runtime/jellyfin-setup-window.test.js dist/main/runtime/jellyfin-client-info-main-deps.test.js dist/main/runtime/jellyfin-cli-main-deps.test.js dist/main/runtime/jellyfin-setup-window-main-deps.test.js`
- `bun run test:fast`
- `bun run docs:build`

## Next Intent (2026-02-21)

- remove `jellyfin.accessToken` config path; add env override (`SUBMINER_JELLYFIN_ACCESS_TOKEN`)
- move `jellyfin.userId` out of config into stored Jellyfin auth session payload
- keep login/logout semantics via setup/CLI commands
- add/update runtime tests first, then implementation + docs

## 2026-02-21 Outcome

- removed `jellyfin.accessToken` + `jellyfin.userId` from config defaults/parsing/docs surface
- added env override precedence in resolver:
  - `SUBMINER_JELLYFIN_ACCESS_TOKEN`
  - `SUBMINER_JELLYFIN_USER_ID` (optional; falls back to stored userId if session exists)
- moved Jellyfin store from token-only to auth-session payload (`accessToken` + `userId`)
- updated login/setup/logout flows to save/clear stored session payload

## 2026-02-21 Files Touched

- `src/core/services/jellyfin-token-store.ts`
- `src/main.ts`
- `src/main/runtime/jellyfin-client-info.ts`
- `src/main/runtime/jellyfin-client-info-main-deps.ts`
- `src/main/runtime/jellyfin-cli-auth.ts`
- `src/main/runtime/jellyfin-cli-main-deps.ts`
- `src/main/runtime/jellyfin-setup-window.ts`
- `src/main/runtime/jellyfin-setup-window-main-deps.ts`
- `src/main/runtime/jellyfin-remote-commands.ts`
- `src/main/runtime/jellyfin-remote-session-lifecycle.ts`
- `src/config/definitions.ts`
- `src/config/service.ts`
- `src/types.ts`
- `src/main/runtime/jellyfin-client-info.test.ts`
- `src/main/runtime/jellyfin-client-info-main-deps.test.ts`
- `src/main/runtime/jellyfin-cli-auth.test.ts`
- `src/main/runtime/jellyfin-cli-main-deps.test.ts`
- `src/main/runtime/jellyfin-setup-window.test.ts`
- `src/main/runtime/jellyfin-setup-window-main-deps.test.ts`
- `docs/configuration.md`
- `docs/jellyfin-integration.md`
- `docs/public/config.example.jsonc`
- `backlog/tasks/task-93 - Remove-Jellyfin-token-userId-from-config-use-env-and-stored-session.md`

## 2026-02-21 Verification

- `bun run tsc --noEmit`
- `bun test src/main/runtime/jellyfin-client-info.test.ts src/main/runtime/jellyfin-cli-auth.test.ts src/main/runtime/jellyfin-setup-window.test.ts src/main/runtime/jellyfin-client-info-main-deps.test.ts src/main/runtime/jellyfin-cli-main-deps.test.ts src/main/runtime/jellyfin-setup-window-main-deps.test.ts src/main/runtime/jellyfin-remote-session-lifecycle.test.ts src/main/runtime/jellyfin-command-dispatch.test.ts src/main/runtime/jellyfin-remote-commands.test.ts src/config/config.test.ts` (73 pass)

## Blockers

- full `bun run build` currently blocked by local disk-full (`ENOSPC`) in workspace output path
