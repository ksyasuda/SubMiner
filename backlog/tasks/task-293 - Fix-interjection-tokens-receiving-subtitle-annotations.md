---
id: TASK-293
title: Fix interjection tokens receiving subtitle annotations
status: In Progress
assignee: []
created_date: '2026-04-25 22:50'
labels:
  - tokenizer
  - bug
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Standalone interjections such as あ should remain hoverable dictionary tokens but must not receive N+1, frequency, JLPT, or known-word subtitle annotation metadata.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 A MeCab 感動詞 token like あ is excluded by the shared subtitle annotation gate.
- [ ] #2 annotateTokens strips N+1/frequency/JLPT/known metadata from the interjection while preserving token lookup fields.
- [ ] #3 Focused tokenizer regression passes.
<!-- AC:END -->
