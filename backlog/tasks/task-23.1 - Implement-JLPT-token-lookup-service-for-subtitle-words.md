---
id: TASK-23.1
title: Implement JLPT token lookup service for subtitle words
status: Done
assignee: []
created_date: '2026-02-13 16:42'
updated_date: '2026-02-18 04:11'
labels: []
dependencies: []
parent_task_id: TASK-23
priority: high
ordinal: 32000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Create a lookup layer that parses/queries the bundled JLPT dictionary file and returns JLPT level for a given token/word. Integrate with subtitle tokenization path with minimal performance overhead.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Service accepts a token/normalized token and returns JLPT level or no-match deterministically.
- [x] #2 Lookup handles expected dictionary format edge cases and unknown tokens without throwing.
- [x] #3 Lookup path is efficient enough for frame-by-frame subtitle updates.
- [x] #4 Tokenizer interaction preserves existing token ordering and positions needed for rendering spans/underlines.
- [x] #5 Behavior on malformed/unsupported dictionary format is documented with fallback semantics.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Lookup service returns JLPT level with deterministic output for test fixtures.
<!-- DOD:END -->
