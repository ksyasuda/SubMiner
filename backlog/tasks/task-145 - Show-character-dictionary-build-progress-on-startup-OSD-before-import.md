---
id: TASK-145
title: Show character dictionary build progress on startup OSD before import
status: Done
assignee: []
created_date: '2026-03-09 11:20'
updated_date: '2026-03-09 11:20'
labels:
  - startup
  - dictionary
  - ux
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/main/runtime/character-dictionary-auto-sync.ts
  - /home/sudacode/projects/japanese/SubMiner/src/main/runtime/startup-osd-sequencer.ts
  - /home/sudacode/projects/japanese/SubMiner/src/main/runtime/character-dictionary-auto-sync.test.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Surface an explicit character-dictionary build phase on startup OSD so there is visible progress between subtitle annotation loading and the later import/upload step when merged dictionary generation is still running.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Auto-sync emits a dedicated in-flight status while merged dictionary generation is running.
- [x] #2 Startup OSD sequencing treats that build phase as progress and can surface it after annotation loading clears.
- [x] #3 Regression coverage verifies the build phase is emitted before import begins.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Added a `building` progress phase before `buildMergedDictionary(...)` and included it in the startup OSD sequencer's buffered progress set. This gives startup a visible dictionary-progress step even when snapshot checking/generation finished too early to still be relevant by the time annotation loading completes.

<!-- SECTION:NOTES:END -->
