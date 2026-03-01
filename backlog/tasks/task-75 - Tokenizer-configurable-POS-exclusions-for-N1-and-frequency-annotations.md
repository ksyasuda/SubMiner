---
id: TASK-75
title: 'Tokenizer: configurable POS exclusions for N+1 and frequency annotations'
status: Done
assignee: []
created_date: '2026-03-01 01:23'
updated_date: '2026-03-01 01:32'
labels: []
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
N+1 and frequency highlighting should ignore non-learning tokens (e.g., particles/auxiliary forms) based on MeCab POS1 tags, while remaining user-configurable.

Problem example: for subtitle phrase containing になれば, the highlighted N+1 target should not be the non-useful inflection/token piece when POS indicates an excluded class.

Implement configurable exclusion defaults with add/remove overrides so users can tune behavior without code changes.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Default exclusion set omits non-useful POS1 classes from both N+1 candidate selection and frequency highlighting.
- [x] #2 Users can add extra POS1 exclusions and remove defaults via config.
- [x] #3 Tokenizer/annotation tests cover default behavior and config add/remove overrides.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented configurable annotation POS exclusions with defaults+add/remove for both MeCab POS1 and POS2, wired to N+1 candidate selection and frequency highlighting. Added POS2 default exclusion (非自立), expanded POS1 defaults for function words, added Yomitan->MeCab enrichment to carry pos2/pos3 metadata, updated config docs/examples, and added regression tests including になれば case.
<!-- SECTION:FINAL_SUMMARY:END -->
