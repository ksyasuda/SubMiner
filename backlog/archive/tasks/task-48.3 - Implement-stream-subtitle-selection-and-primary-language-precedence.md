---
id: TASK-48.3
title: Implement stream subtitle selection and primary-language precedence
status: To Do
assignee: []
created_date: '2026-02-14 06:03'
updated_date: '2026-02-14 08:19'
labels:
  - stream
  - subtitles
dependencies: []
parent_task_id: TASK-48
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Handle subtitle track selection for stream playback so Japanese (or configured primary language) subtitle behavior is correctly applied when stream metadata includes or omits JP tracks.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Use stream metadata to choose and mark the configured primary language subtitle as active when available.
- [ ] #2 If no matching primary language track exists in stream metadata, keep previous fallback behavior only for non-streaming mode.
- [ ] #3 When no Japanese track exists and config primary is different, explicitly set configured primary as primary track for streaming flow.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Superseded by TASK-51. Stream subtitle language precedence in streaming mode deferred with full design revisit.
<!-- SECTION:NOTES:END -->
