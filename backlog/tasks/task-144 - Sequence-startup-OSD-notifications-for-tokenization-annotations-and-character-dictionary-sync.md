---
id: TASK-144
title: Sequence startup OSD notifications for tokenization, annotations, and character dictionary sync
status: Done
assignee: []
created_date: '2026-03-09 10:40'
updated_date: '2026-03-09 10:40'
labels:
  - startup
  - overlay
  - ux
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/main.ts
  - /home/sudacode/projects/japanese/SubMiner/src/main/runtime/startup-osd-sequencer.ts
  - /home/sudacode/projects/japanese/SubMiner/src/main/runtime/subtitle-tokenization-main-deps.ts
  - /home/sudacode/projects/japanese/SubMiner/src/main/runtime/character-dictionary-auto-sync-notifications.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Keep startup OSD progress ordered. While tokenization is still pending, only show the tokenization loading message. After tokenization becomes ready, show annotation loading if annotation warmup still remains. Only surface character dictionary auto-sync progress after annotation loading clears, and only if the dictionary work is still active.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Character dictionary progress stays hidden while tokenization startup loading is still active.
- [x] #2 Annotation loading OSD appears after tokenization readiness and before any later character dictionary progress.
- [x] #3 Regression coverage verifies buffered dictionary progress/failure ordering during startup.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Added a small startup OSD sequencer in main runtime. Annotation warmup OSD now flows through that sequencer, and character dictionary sync notifications buffer until tokenization plus annotation loading clear. Buffered `ready` updates are dropped if dictionary progress finished before it ever became visible, while buffered failures still surface after annotation loading completes.

<!-- SECTION:NOTES:END -->
