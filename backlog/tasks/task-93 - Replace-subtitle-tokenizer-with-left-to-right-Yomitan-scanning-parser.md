---
id: TASK-93
title: Replace subtitle tokenizer with left-to-right Yomitan scanning parser
status: Done
assignee: []
created_date: '2026-03-06 09:02'
updated_date: '2026-03-16 05:13'
labels:
  - tokenizer
  - yomitan
  - refactor
dependencies: []
priority: high
ordinal: 85500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Replace the current parseText candidate-selection tokenizer with a GSM-style left-to-right Yomitan scanning tokenizer for all subtitles. Preserve downstream token contracts for rendering, JLPT/frequency/N+1 annotation, and MeCab enrichment while improving full-term matching for names and katakana compounds.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Subtitle tokenization uses a left-to-right Yomitan scanning strategy instead of parseText candidate selection.
- [x] #2 Token surfaces, readings, headwords, and offsets remain compatible with existing renderer and annotation stages.
- [x] #3 Known problematic name cases such as カズマ and バニール resolve to full-token dictionary matches when Yomitan can match them.
- [x] #4 Regression tests cover left-to-right exact-match scanning, unmatched text handling, and downstream tokenizeSubtitle integration.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Replaced the live subtitle tokenization path with a left-to-right Yomitan `termsFind` scanner that greedily advances through the normalized subtitle text, preserving downstream `MergedToken` contracts for renderer, MeCab enrichment, JLPT, frequency, and N+1 annotation. Added runtime and integration coverage for exact-match scanning plus name cases like カズマ and kept compatibility fallback handling for older mocked parseText-style test payloads.
<!-- SECTION:FINAL_SUMMARY:END -->
