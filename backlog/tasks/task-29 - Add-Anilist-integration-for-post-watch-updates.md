---
id: TASK-29
title: Add Anilist integration for post-watch updates
status: Done
assignee: []
created_date: '2026-02-13 17:57'
updated_date: '2026-02-18 04:11'
labels:
  - anilist
  - anime
  - integration
  - electron
  - api
dependencies: []
priority: medium
ordinal: 16000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Add Anilist integration so the app can update user anime progress after watching, by porting the core functionality of `AzuredBlue/mpv-anilist-updater` into the Electron app. The initial implementation should focus on reliable sync of watch status/progress and be structured to support future Anilist features beyond updates.

Requirements:

- Port the core behavior from `AzuredBlue/mpv-anilist-updater` needed for post-watch syncing into the Electron architecture.
- Authenticate and persist Anilist credentials securely in the desktop environment.
- Identify anime/media item and track watched status/progress based on existing video/session data.
- Trigger updates to Anilist after watch milestones/session completion.
- Queue and retry updates safely when offline or API errors occur.
- Avoid blocking playback/mining operations while syncing.
- Design Anilist integration as a dedicated, testable module/service so additional Anilist features can be added later (e.g., status actions, favorites, episode tracking enhancements).
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Application can authenticate with Anilist and store tokens securely for desktop user sessions.
- [x] #2 Anilist update flow from existing local watch/session metadata can update animes watched progress after watching.
- [x] #3 Core functionality equivalent to mpv-anilist-updater is implemented for this use case (progress/status sync) inside the Electron app.
- [x] #4 A background/in-process queue handles transient API failures and retries without losing updates.
- [x] #5 Sync updates are non-blocking and do not degrade normal playback/mining behavior.
- [x] #6 Anilist integration code is modularized to allow future feature additions without major refactor (clear service boundaries/interfaces).
- [x] #7 Error states and duplicate/duplicate-inconsistent updates are handled deterministically (idempotent where practical).
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Completed child tasks TASK-29.1 and TASK-29.2: secure token persistence/fallback and persistent retry queue with backoff/dead-letter are now implemented.

Implemented AniList control surfaces across CLI and IPC. CLI: added `--anilist-status`, `--anilist-logout`, `--anilist-setup`, `--anilist-retry-queue` parsing/help/dispatch/runtime wiring in `src/cli/args.ts`, `src/cli/help.ts`, `src/core/services/cli-command.ts`, `src/main/cli-runtime.ts`, and `src/main.ts`. IPC: added `anilist:get-status`, `anilist:clear-token`, `anilist:open-setup`, `anilist:get-queue-status`, `anilist:retry-now` handlers and dependency wiring in `src/core/services/ipc.ts`, `src/main/dependencies.ts`, and `src/main.ts`. Added retry queue tests in `src/core/services/anilist/anilist-update-queue.test.ts`; added token-store tests in `src/core/services/anilist/anilist-token-store.test.ts` with runtime guard for environments without Electron safeStorage; updated `package.json` test list and AniList command/channel docs in `docs/configuration.md`.

Validation run: `pnpm run test:fast` passed (config + core dist suite). New AniList queue tests run in core suite; token-store tests are auto-skipped in plain Node environments where Electron `safeStorage` is unavailable (still active when safeStorage exists).

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Completed AniList post-watch integration in the Electron app with secure token lifecycle, post-watch progress sync, and persistent retry/backoff queue with dead-letter behavior. Added user-facing control surfaces via CLI (`--anilist-status`, `--anilist-logout`, `--anilist-setup`, `--anilist-retry-queue`) and IPC (`anilist:get-status`, `anilist:clear-token`, `anilist:open-setup`, `anilist:get-queue-status`, `anilist:retry-now`), plus docs and test coverage; validated via `pnpm run test:fast`.

<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done

<!-- DOD:BEGIN -->

- [x] #1 Core Anilist service module exists and is wired into application flow for post-watch updates.
- [x] #2 OAuth/token lifecycle is implemented with safe local persistence and revocation/logout behavior.
- [x] #3 A retry/backoff and dead-letter strategy for failed syncs is implemented.
- [x] #4 User-visible settings/docs explain how to connect/manage Anilist and what data is synced.
- [x] #5 At least smoke/integration coverage (or validated test plan) for mapping and sync flow is in place.
<!-- DOD:END -->
