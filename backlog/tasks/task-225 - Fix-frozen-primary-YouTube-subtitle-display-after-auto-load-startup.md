---
id: TASK-225
title: Fix frozen primary YouTube subtitle display after auto-load startup
status: Done
assignee: []
created_date: '2026-03-23 20:07'
updated_date: '2026-03-23 20:15'
labels:
  - bug
  - youtube
  - subtitles
dependencies:
  - TASK-224
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
After the new YouTube auto-load startup flow, the primary subtitle overlay can stay stuck on an older line while the subtitle sidebar continues advancing. Investigate startup suppression / subtitle refresh timing and restore live primary overlay updates after auto-loaded subtitles are injected.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 When YouTube auto-load succeeds, the visible primary subtitle continues advancing after playback resumes.
- [x] #2 Startup suppression does not leave the primary subtitle display stuck on a stale line.
- [x] #3 A regression test covers the startup path that previously froze the visible primary subtitle while sidebar timing continued advancing.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Root cause: applyStartupState seeded youtubePlaybackFlowPending from initialArgs.youtubePlay, and runYoutubePlaybackFlowMain restored that preexisting true value after startup auto-load. Result: primary subtitle events stayed suppressed for startup-launched YouTube playback while sidebar timing still advanced.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Stopped pre-seeding youtubePlaybackFlowPending from startup CLI args so only the actual YouTube playback bootstrap window suppresses subtitle events. Added a regression test covering startup YouTube args and re-ran targeted YouTube/runtime subtitle tests plus typecheck.
<!-- SECTION:FINAL_SUMMARY:END -->
