---
id: TASK-213
title: Show character dictionary progress during paused startup waits
status: Done
assignee: []
created_date: '2026-03-20 08:59'
updated_date: '2026-03-23 03:22'
labels:
  - bug
  - ux
  - dictionary
  - startup
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/main/runtime/startup-osd-sequencer.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/main/runtime/character-dictionary-auto-sync-notifications.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/main.ts
priority: medium
ordinal: 141500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
During startup on mpv auto-start, character dictionary regeneration/update can be active while playback remains paused. The current startup OSD sequencer buffers dictionary progress behind annotation-loading OSD, which leaves the user with no visible dictionary-specific progress while the pause is active. Adjust the startup OSD sequencing so dictionary progress can surface once tokenization is ready during the paused startup window, without regressing later ready/failure handling.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 When tokenization is ready during startup, later character dictionary progress updates are shown on OSD even if annotation-loading state is still active.
- [ ] #2 Startup OSD completion/failure behavior for character dictionary sync remains coherent after the new progress ordering.
- [ ] #3 Regression coverage exercises the paused startup sequencing for dictionary progress.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-03-20: Confirmed issue is broader than OSD-only. Paused-startup OSD fixes remain relevant, but current user report also points at a regression in non-blocking startup playback release (tracked in TASK-143).

2026-03-20: OSD sequencing fix remains in local patch alongside TASK-143 regression fix. Covered by startup-osd-sequencer tests; pending installed-app/mpv validation before task finalization.
<!-- SECTION:NOTES:END -->
