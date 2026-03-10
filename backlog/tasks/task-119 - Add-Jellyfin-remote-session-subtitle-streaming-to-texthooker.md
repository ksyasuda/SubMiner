---
id: TASK-119
title: Add Jellyfin remote-session subtitle streaming to texthooker
status: To Do
assignee: []
created_date: '2026-03-08 03:46'
labels:
  - jellyfin
  - texthooker
  - subtitle
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/main/runtime/jellyfin-remote-commands.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/core/services/jellyfin.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/subtitle-processing-controller.ts
  - 'https://api.jellyfin.org/'
documentation:
  - 'https://api.jellyfin.org/'
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Allow SubMiner to follow subtitles from a separate Jellyfin client session, such as a TV app, without requiring local mpv playback. The feature should fetch the active subtitle stream from Jellyfin, map the remote playback position to subtitle cues, and feed the existing subtitle tokenization plus annotated texthooker websocket pipeline so texthooker-only mode can be used while watching on another device.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 User can target a remote Jellyfin session and stream its current subtitle cue into SubMiner's existing subtitle-processing pipeline without launching local Jellyfin playback in mpv.
- [ ] #2 Texthooker-only mode can display subtitle updates from the tracked remote Jellyfin session through the existing annotation websocket feed.
- [ ] #3 Remote session changes are handled safely: item changes, subtitle-track changes, pause/seek/stop, and session disconnects clear or refresh subtitle state without crashing.
- [ ] #4 The feature degrades clearly when the remote session has no usable text subtitle stream or uses an unsupported subtitle format.
- [ ] #5 Automated tests cover session tracking, subtitle cue selection, and feed integration; user-facing docs/config docs are updated.
<!-- AC:END -->
