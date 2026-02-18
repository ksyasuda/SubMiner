---
id: TASK-60
title: Remove hard-coded particle term exclusions from frequency lookup
status: Done
assignee: []
created_date: '2026-02-16 22:20'
updated_date: '2026-02-18 04:11'
labels: []
dependencies: []
ordinal: 25000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Update tokenizer frequency filtering to rely on MeCab POS information instead of a hard-coded set of particle surface forms.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 `FREQUENCY_EXCLUDED_PARTICLES` hard-coded term list is removed.
- [x] #2 Frequency exclusion for particles/auxiliaries is driven by POS metadata.
- [x] #3 Tokenizer tests cover POS-driven exclusion behavior.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Removed hard-coded particle surface exclusions (`FREQUENCY_EXCLUDED_PARTICLES`) from tokenizer frequency logic. Frequency skip now relies on POS metadata only: `partOfSpeech` (`particle`/`bound_auxiliary`) and MeCab-enriched `pos1` (`助詞`/`助動詞`) for Yomitan tokens. Added tokenizer test `tokenizeSubtitleService skips frequency rank when Yomitan token is enriched as particle by mecab pos1` to validate POS-driven exclusion.

<!-- SECTION:FINAL_SUMMARY:END -->
