---
id: TASK-31
title: Add optional Jellyfin integration with basic streaming/ playback features
status: To Do
assignee: []
created_date: '2026-02-13 18:38'
labels: []
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Implement optional Jellyfin integration so SubMiner can act as a lightweight Jellyfin client similar to jellyfin-mpv-shim. The feature should support connecting to Jellyfin servers, browsing playable media, and launching playback through SubMiner, including direct play when possible and transparent transcoding when required.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Add a configurable Jellyfin integration path that can be enabled/disabled without impacting core non-Jellyfin functionality.
- [ ] #2 Support authenticating against a user-selected Jellyfin server (server URL + credentials/token) and securely storing/reusing connection settings.
- [ ] #3 Allow discovery or manual selection of movies/tv shows/music libraries and playback items from the connected Jellyfin server.
- [ ] #4 Enable playback from Jellyfin items via existing player pipeline with a dedicated selection/launch flow.
- [ ] #5 Honor Jellyfin playback options so direct play is attempted first when media/profiles are compatible.
- [ ] #6 Fall back to Jellyfin-managed transcoding when direct play is not possible, passing required transcode parameters to the player.
- [ ] #7 Preserve useful Jellyfin metadata/features during playback: title/season/episode, subtitles, audio track selection, and playback resume markers where available.
- [ ] #8 Add handling for common failure modes (invalid credentials, token expiry, server offline, transcoding/stream errors) with user-visible status/errors.
- [ ] #9 Document setup and limitations (what works vs what is optional) in project documentation, and add tests or mocks that validate key integration logic and settings handling.
<!-- AC:END -->
