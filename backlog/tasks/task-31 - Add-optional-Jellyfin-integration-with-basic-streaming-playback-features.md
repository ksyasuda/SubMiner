---
id: TASK-31
title: Add optional Jellyfin integration with basic streaming/ playback features
status: Done
assignee: []
created_date: '2026-02-13 18:38'
updated_date: '2026-02-18 04:11'
labels: []
dependencies: []
references:
  - TASK-64
  - docs/plans/2026-02-17-jellyfin-cast-remote-playback.md
ordinal: 57000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Implement optional Jellyfin integration so SubMiner can act as a lightweight Jellyfin client similar to jellyfin-mpv-shim. The feature should support connecting to Jellyfin servers, browsing playable media, and launching playback through SubMiner, including direct play when possible and transparent transcoding when required.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Add a configurable Jellyfin integration path that can be enabled/disabled without impacting core non-Jellyfin functionality.
- [x] #2 Support authenticating against a user-selected Jellyfin server (server URL + credentials/token) and securely storing/reusing connection settings.
- [x] #3 Allow discovery or manual selection of movies/tv shows/music libraries and playback items from the connected Jellyfin server.
- [x] #4 Enable playback from Jellyfin items via existing player pipeline with a dedicated selection/launch flow.
- [x] #5 Honor Jellyfin playback options so direct play is attempted first when media/profiles are compatible.
- [x] #6 Fall back to Jellyfin-managed transcoding when direct play is not possible, passing required transcode parameters to the player.
- [x] #7 Preserve useful Jellyfin metadata/features during playback: title/season/episode, subtitles, audio track selection, and playback resume markers where available.
- [x] #8 Add handling for common failure modes (invalid credentials, token expiry, server offline, transcoding/stream errors) with user-visible status/errors.
- [x] #9 Document setup and limitations (what works vs what is optional) in project documentation, and add tests or mocks that validate key integration logic and settings handling.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Status snapshot (2026-02-18): TASK-31 is mostly complete and now tracks remaining closure work only for #2 and #3.

Completed acceptance criteria and evidence:

- #1 Optional/disabled Jellyfin integration boundary verified.
  - Added tests in `src/core/services/app-ready.test.ts`, `src/core/services/cli-command.test.ts`, `src/core/services/startup-bootstrap.test.ts`, `src/core/services/jellyfin-remote.test.ts`, and `src/config/config.test.ts` to prove disabled paths do not impact core non-Jellyfin functionality and that Jellyfin side effects are gated.
- #4 Jellyfin playback launch through existing pipeline verified.
- #5 Direct-play preference behavior verified.
  - `resolvePlaybackPlan` chooses direct when compatible/preferred and switches away from direct when preference/compatibility disallows it.
- #6 Transcode fallback behavior verified.
  - `resolvePlaybackPlan` falls back to transcode and preserves required params (`api_key`, stream indexes, resume ticks, codec params).
- #7 Metadata/subtitle/audio/resume parity (within current scope) verified.
  - Added tests proving episode title formatting, stream selection propagation, resume marker handling, and subtitle-track fallback behavior.
- #8 Failure-mode handling and user-visible error surfacing verified.
  - Added tests for invalid credentials (401), expired/invalid token auth failures (403), non-OK server responses, no playable source / no stream path, and CLI OSD error surfacing (`Jellyfin command failed: ...`).
- #9 Docs + key integration tests/mocks completed.

Key verification runs (all passing):

- `pnpm run build`
- `node --test dist/core/services/app-ready.test.js dist/core/services/cli-command.test.js dist/core/services/startup-bootstrap.test.js dist/core/services/jellyfin-remote.test.js dist/config/config.test.js`
- `node --test dist/core/services/jellyfin.test.js dist/core/services/cli-command.test.js`
- `pnpm run test:fast`

Open acceptance criteria (remaining work):

- #2 Authentication/settings persistence hardening and explicit lifecycle validation:
  1. login -> persist -> restart -> token reuse verification
  2. token-expiry re-auth/recovery path verification
  3. document storage guarantees/edge cases
- #3 Library discovery/manual selection UX closure across intended media scope:
  1. explicit verification for movies/TV/music discovery and selection paths
  2. document any intentionally out-of-scope media types/flows

Task relationship:

- TASK-64 remains a focused implementation slice under this epic and provides foundational cast/remote playback work referenced by this task.
<!-- SECTION:NOTES:END -->
