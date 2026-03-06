---
id: TASK-92
title: Fix merged Yomitan headword selection for katakana subtitle tokens
status: Done
assignee: []
created_date: '2026-03-06 08:43'
updated_date: '2026-03-06 08:43'
labels:
  - bug
  - tokenizer
  - yomitan
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Tokenizer/parser-selection bug: when a scanning-parser line is merged from multiple segments, the merged token currently keeps the first segment headword even if a later segment provides the full dictionary-backed term. This truncates katakana names such as バニール to バニ in the lookup payload and prevents correct dictionary matching. Also align kana classification so the prolonged sound mark is treated as kana in tokenizer heuristics.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Merged scanning-parser tokens prefer a full cross-segment headword when one segment expands to the full term.
- [x] #2 Standalone later segment headwords do not override the primary token headword in normal content-word + auxiliary merges.
- [x] #3 Katakana prolonged sound mark is treated as kana in tokenizer heuristics.
- [x] #4 Regression tests cover the merged katakana headword case.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Adjusted merged scanning-parser headword selection so later segments only override the first headword when they provide an expanded cross-segment dictionary term, which fixes truncated katakana lookups like バニール -> バニ. Also updated kana classification to include the katakana prolonged sound mark and added regression coverage for both the expanded-headword case and the normal content-word-plus-auxiliary case.
<!-- SECTION:FINAL_SUMMARY:END -->
