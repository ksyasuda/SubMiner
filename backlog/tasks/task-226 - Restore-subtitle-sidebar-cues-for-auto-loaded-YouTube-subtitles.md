---
id: TASK-226
title: Restore subtitle sidebar cues for auto-loaded YouTube subtitles
status: Done
assignee: []
created_date: '2026-03-23 20:21'
updated_date: '2026-03-23 20:25'
labels:
  - bug
  - youtube
  - subtitle-sidebar
dependencies:
  - TASK-224
  - TASK-225
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
After fixing startup subtitle event suppression, the primary subtitle overlay updates for auto-loaded YouTube playback but the subtitle sidebar reports no parsed subtitle cues available. Investigate parsed subtitle source registration / refresh for auto-loaded YouTube subtitle files and restore sidebar cue population.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 When YouTube auto-load succeeds, the subtitle sidebar receives parsed cues for the active primary subtitle source.
- [x] #2 Auto-loaded YouTube subtitle source changes refresh the sidebar snapshot without requiring manual picker interaction.
- [x] #3 A regression test covers the startup auto-load path where live primary subtitles render but sidebar cues remain empty.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Root cause: successful YouTube auto-load refreshed visible primary subtitle state, but did not explicitly initialize parsed subtitle cues from the resolved downloaded primary subtitle file. Sidebar cue population depended on later mpv source rediscovery, which could leave snapshots empty.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added a regression test for YouTube auto-load sidebar cue refresh and wired the YouTube subtitle flow to explicitly refresh parsed subtitle cues from the resolved primary subtitle path after a successful load. Verified with targeted YouTube/sidebar/runtime tests plus typecheck.
<!-- SECTION:FINAL_SUMMARY:END -->
