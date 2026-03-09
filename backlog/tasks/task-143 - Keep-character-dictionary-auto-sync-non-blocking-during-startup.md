---
id: TASK-143
title: Keep character dictionary auto-sync non-blocking during startup
status: Done
assignee: []
created_date: '2026-03-09 01:45'
updated_date: '2026-03-09 01:45'
labels:
  - dictionary
  - startup
  - performance
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/main.ts
  - /home/sudacode/projects/japanese/SubMiner/src/main/runtime/character-dictionary-auto-sync.ts
  - /home/sudacode/projects/japanese/SubMiner/src/main/runtime/current-media-tokenization-gate.ts
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Keep character dictionary auto-sync running in parallel during startup without delaying playback. Only tokenization readiness should gate playback; character dictionary import/settings updates should wait until tokenization is already ready and then refresh annotations afterward.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Character dictionary snapshot/build work can run immediately during startup.
- [x] #2 Yomitan dictionary mutation work waits until current-media tokenization is ready.
- [x] #3 Regression coverage verifies auto-sync builds before the gate and only mutates Yomitan after the gate resolves.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Added a small current-media tokenization gate in main runtime. Media changes reset the gate, the first tokenization-ready event marks it ready, and auto-sync now waits on that gate only before Yomitan dictionary inspection/import/settings updates. Snapshot generation and merged ZIP build still run immediately in parallel.

<!-- SECTION:NOTES:END -->
