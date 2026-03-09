---
id: TASK-141
title: Refresh current subtitle after character dictionary sync completes
status: Done
assignee: []
created_date: '2026-03-09 00:00'
updated_date: '2026-03-09 00:55'
labels:
  - dictionary
  - overlay
  - bug
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/main/runtime/character-dictionary-auto-sync.ts
  - /home/sudacode/projects/japanese/SubMiner/src/main.ts
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
When character dictionary auto-sync finishes after startup tokenization, invalidate cached subtitle tokenization and refresh the current subtitle so character-name highlighting catches up without waiting for the next subtitle line.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Successful character dictionary sync exposes a completion hook for main runtime follow-up.
- [x] #2 Main runtime clears Yomitan parser caches and refreshes the current subtitle after sync completion.
- [x] #3 Regression coverage verifies the sync completion callback fires on successful sync.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Observed on Bunny Girl Senpai startup: autoplay/tokenization became ready around 8s, but snapshot/import/state write completed roughly 31s after launch, leaving the current subtitle tokenized without the newly imported character dictionary. Fixed by adding an auto-sync completion hook that clears parser caches and refreshes the current subtitle.
<!-- SECTION:NOTES:END -->
