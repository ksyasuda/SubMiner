---
id: TASK-29
title: Add Anilist integration for post-watch updates
status: In Progress
assignee: []
created_date: '2026-02-13 17:57'
updated_date: '2026-02-17 04:19'
labels:
  - anilist
  - anime
  - integration
  - electron
  - api
dependencies: []
priority: medium
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
- [ ] #1 Application can authenticate with Anilist and store tokens securely for desktop user sessions.
- [ ] #2 Anilist update flow from existing local watch/session metadata can update animes watched progress after watching.
- [ ] #3 Core functionality equivalent to mpv-anilist-updater is implemented for this use case (progress/status sync) inside the Electron app.
- [ ] #4 A background/in-process queue handles transient API failures and retries without losing updates.
- [ ] #5 Sync updates are non-blocking and do not degrade normal playback/mining behavior.
- [ ] #6 Anilist integration code is modularized to allow future feature additions without major refactor (clear service boundaries/interfaces).
- [ ] #7 Error states and duplicate/duplicate-inconsistent updates are handled deterministically (idempotent where practical).
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Completed child tasks TASK-29.1 and TASK-29.2: secure token persistence/fallback and persistent retry queue with backoff/dead-letter are now implemented.
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Core Anilist service module exists and is wired into application flow for post-watch updates.
- [ ] #2 OAuth/token lifecycle is implemented with safe local persistence and revocation/logout behavior.
- [ ] #3 A retry/backoff and dead-letter strategy for failed syncs is implemented.
- [ ] #4 User-visible settings/docs explain how to connect/manage Anilist and what data is synced.
- [ ] #5 At least smoke/integration coverage (or validated test plan) for mapping and sync flow is in place.
<!-- DOD:END -->
