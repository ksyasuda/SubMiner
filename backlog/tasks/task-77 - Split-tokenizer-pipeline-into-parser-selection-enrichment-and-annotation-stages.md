---
id: TASK-77
title: Split tokenizer pipeline into parser selection enrichment and annotation stages
status: To Do
assignee: []
created_date: '2026-02-18 11:43'
updated_date: '2026-02-18 11:43'
labels:
  - tokenizer
  - subtitles
  - refactor
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`src/core/services/tokenizer.ts` mixes parser-window lifecycle, candidate scoring, MeCab enrichment, and post-annotation (known/JLPT/frequency/N+1). This task decomposes the tokenizer into explicit pipeline stages with stable contracts.
<!-- SECTION:DESCRIPTION:END -->

## Suggestions

<!-- SECTION:SUGGESTIONS:BEGIN -->
- Use a stage pipeline contract (`TokenizationInput -> StageOutput`) to isolate concerns.
- Isolate Yomitan parser-window lifecycle into a dedicated runtime adapter module.
- Keep heuristics/test fixtures in one place to make scoring behavior reviewable.
<!-- SECTION:SUGGESTIONS:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Define stage interfaces: source parsing, candidate selection, POS enrichment, semantic annotation.
2. Extract parser runtime adapter (`ensure parser window`, `execute parse`) from annotation logic.
3. Extract candidate scoring/selection into pure module with fixture-based tests.
4. Extract annotation passes (known-word, frequency, JLPT, N+1) into composable functions.
5. Add regression tests with fixed subtitle inputs to lock behavior.
6. Update tokenizer architecture notes in docs.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Tokenizer code is split into explicit stages with narrow interfaces
- [ ] #2 Candidate selection logic is pure + directly testable
- [ ] #3 Parser lifecycle concerns are separated from annotation passes
- [ ] #4 Existing tokenization behavior preserved in regression tests
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Tokenizer-related test suites pass
- [ ] #2 New stage-level tests exist for scoring and annotation
<!-- DOD:END -->

