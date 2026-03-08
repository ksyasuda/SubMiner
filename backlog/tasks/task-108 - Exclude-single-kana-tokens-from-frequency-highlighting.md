---
id: TASK-108
title: 'Exclude single kana tokens from frequency highlighting'
status: Done
assignee: []
created_date: '2026-03-07 01:18'
updated_date: '2026-03-07 01:22'
labels: []
dependencies: []
priority: medium
ordinal: 9008
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Suppress frequency highlighting for single-character hiragana or katakana tokens. Scope is frequency-only: known/N+1/JLPT behavior stays unchanged.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Single-character hiragana tokens do not retain `frequencyRank`.
- [x] #2 Single-character katakana tokens do not retain `frequencyRank`.
- [x] #3 Regression coverage exists at annotation-stage and tokenizer levels.

<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Added a frequency-only suppression rule for single-character kana tokens based on token `surface`, so bogus merged fragments like `た` and standalone one-character kana no longer keep `frequencyRank`. Regression coverage now exists both in the annotation stage and in the tokenizer path, while multi-character tokens and N+1/JLPT behavior remain unchanged.

Verification:

- `bun test src/core/services/tokenizer/annotation-stage.test.ts --timeout 20000`
- `bun test src/core/services/tokenizer.test.ts --timeout 20000`

<!-- SECTION:FINAL_SUMMARY:END -->
