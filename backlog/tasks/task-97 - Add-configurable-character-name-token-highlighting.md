---
id: TASK-97
title: Add configurable character-name token highlighting
status: Done
assignee: []
created_date: '2026-03-06 10:15'
updated_date: '2026-03-06 10:15'
labels:
  - subtitle
  - dictionary
  - renderer
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/core/services/tokenizer.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/core/services/tokenizer/yomitan-parser-runtime.ts
  - /home/sudacode/projects/japanese/SubMiner/src/renderer/subtitle-render.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Color subtitle tokens that match entries from the SubMiner character dictionary, with a configurable default color and a config toggle that disables both rendering and name-match detection work.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Tokens matched from the SubMiner character dictionary receive dedicated renderer styling.
- [x] #2 `subtitleStyle.nameMatchEnabled` disables name-match detection work when false.
- [x] #3 `subtitleStyle.nameMatchColor` overrides the default `#f5bde6`.
- [x] #4 Regression coverage verifies config parsing, tokenizer propagation, scanner gating, and renderer class/CSS behavior.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added configurable character-name token highlighting with default color `#f5bde6` and config gate `subtitleStyle.nameMatchEnabled`. When enabled, left-to-right Yomitan scanning tags tokens whose winning dictionary entry comes from the SubMiner character dictionary; when disabled, the tokenizer skips that metadata work and the renderer suppresses name-match styling. Added focused regression tests for config parsing, main-deps wiring, Yomitan scan gating, token propagation, renderer classes, and CSS behavior.
<!-- SECTION:FINAL_SUMMARY:END -->
