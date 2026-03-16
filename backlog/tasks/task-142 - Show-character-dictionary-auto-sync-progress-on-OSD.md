---
id: TASK-142
title: Show character dictionary auto-sync progress on OSD
status: Done
assignee: []
created_date: '2026-03-09 01:10'
updated_date: '2026-03-16 05:13'
labels:
  - dictionary
  - overlay
  - ux
dependencies: []
references:
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/runtime/character-dictionary-auto-sync.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/runtime/character-dictionary-auto-sync-notifications.ts
  - /home/sudacode/projects/japanese/SubMiner/src/main.ts
priority: medium
ordinal: 40500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
When character dictionary auto-sync runs for a newly opened anime, surface progress so users know why character-name lookup/highlighting is temporarily unavailable via the mpv OSD without desktop notification popups.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Character dictionary auto-sync emits progress events for syncing, importing, ready, and failure states.
- [x] #2 Main runtime routes those progress events through OSD notifications without desktop notifications.
- [x] #3 Regression coverage verifies progress events and notification routing behavior.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
OSD now shows auto-sync phase changes while the dictionary updates. Desktop notifications were removed for this path to avoid startup popup spam.
<!-- SECTION:NOTES:END -->
