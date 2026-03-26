---
id: TASK-91
title: >-
  Keep unsupported subtitle characters visible while excluding them from token
  hover
status: Done
assignee: []
created_date: '2026-03-06 08:29'
updated_date: '2026-03-16 05:13'
labels:
  - bug
  - tokenizer
  - renderer
dependencies: []
priority: medium
ordinal: 88500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Tokenizer/rendering bug: symbols and other unsupported characters with no lookup result are removed from the rendered subtitle line after tokenization, causing the displayed line to diverge from the source subtitle text. Update rendering so unsupported spans remain visible as plain text but are not tokenized/hoverable, and add regression coverage.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Subtitle rendering preserves unsupported symbols and special characters from the original line.
- [x] #2 Unsupported symbols and special characters do not create interactive token hover targets.
- [x] #3 Regression tests cover a mixed line containing tokenizable text plus unsupported characters.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Updated tokenized subtitle rendering to preserve unsupported punctuation and symbol spans as plain text while keeping only matched tokens interactive. Added renderer and alignment regression coverage for mixed lines so hover offsets stay correct after non-tokenizable characters remain visible.
<!-- SECTION:FINAL_SUMMARY:END -->
