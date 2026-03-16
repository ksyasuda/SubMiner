---
id: TASK-145
title: Show checking and generation OSD for character dictionary auto-sync
status: Done
assignee: []
created_date: '2026-03-09 11:20'
updated_date: '2026-03-16 05:13'
labels:
  - dictionary
  - overlay
  - ux
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/main.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/runtime/character-dictionary-auto-sync.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/runtime/startup-osd-sequencer.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/character-dictionary-runtime.ts
priority: medium
ordinal: 35500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Surface an immediate startup OSD that the character dictionary is being checked, and show a distinct generating message only when the current AniList media actually needs a fresh snapshot build instead of reusing a cached one.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Auto-sync emits a `checking` progress event before snapshot resolution completes.
- [x] #2 Auto-sync emits `generating` only for snapshot cache misses and keeps `updating`/`importing` as later phases.
- [x] #3 Startup OSD sequencing still prioritizes tokenization then annotation loading before buffered dictionary progress.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Character dictionary auto-sync now emits `Checking character dictionary...` as soon as the AniList media is resolved, then emits `Generating character dictionary...` only when the snapshot layer misses and a real rebuild begins. Cached snapshots skip the generating phase and continue straight into the later update/import flow.

Wired those progress callbacks through the character-dictionary runtime boundary, updated the startup OSD sequencer to treat checking/generating as dictionary-progress phases with the same tokenization and annotation precedence, and added regression coverage for cache-hit vs cache-miss behavior plus buffered startup ordering.
<!-- SECTION:FINAL_SUMMARY:END -->
