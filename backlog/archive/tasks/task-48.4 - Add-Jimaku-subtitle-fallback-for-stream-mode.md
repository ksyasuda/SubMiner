---
id: TASK-48.4
title: Add Jimaku subtitle fallback for stream mode
status: To Do
assignee: []
created_date: '2026-02-14 06:03'
updated_date: '2026-02-14 08:19'
labels:
  - stream
  - jimaku
  - subtitles
dependencies: []
parent_task_id: TASK-48
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
When a resolved stream lacks JP/primary-language tracks, fetch subtitles from Jimaku and inject them for playback, overriding non-primary subtitle defaults in streaming mode according to config.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Detect missing primary subtitle from stream metadata and trigger Jimaku lookup for matching episode.
- [ ] #2 Load fetched Jimaku subtitles into playback pipeline and mark them as the primary subtitle track.
- [ ] #3 Fallback is only used in streaming mode and should not alter subtitle behavior outside streaming.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Superseded by TASK-51. Jimaku fallback for streams deferred along with entire streaming flow.
<!-- SECTION:NOTES:END -->
