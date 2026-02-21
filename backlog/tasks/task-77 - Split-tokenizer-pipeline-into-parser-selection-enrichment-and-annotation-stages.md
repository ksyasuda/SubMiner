---
id: TASK-77
title: >-
  Split tokenizer pipeline into parser selection enrichment and annotation
  stages
status: Done
assignee:
  - '@opencode'
created_date: '2026-02-18 11:43'
updated_date: '2026-02-21 23:47'
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
- [x] #1 Tokenizer code is split into explicit stages with narrow interfaces
- [x] #2 Candidate selection logic is pure + directly testable
- [x] #3 Parser lifecycle concerns are separated from annotation passes
- [x] #4 Existing tokenization behavior preserved in regression tests
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Extract pure parser-selection stage from `src/core/services/tokenizer.ts` into `src/core/services/tokenizer/parser-selection-stage.ts` (parse-result mapping + candidate scoring/selection) and add direct stage tests for source preference/tie-break scoring.
2. Extract MeCab POS1 enrichment stage into `src/core/services/tokenizer/parser-enrichment-stage.ts` with direct tests for overlap and surface-sequence fallback behavior.
3. Extract annotation stage into `src/core/services/tokenizer/annotation-stage.ts` to handle known-word/frequency/JLPT/N+1 passes behind a narrow API, with new stage-level tests.
4. Separate parser window/runtime lifecycle into `src/core/services/tokenizer/yomitan-parser-runtime.ts`, keep `tokenizer.ts` as thin orchestrator, run tokenizer + core src/dist gates, then finalize TASK-77 AC/DoD evidence in Backlog MCP.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-02-21: started execution pass in current session; loaded Backlog context and tokenizer module/tests before drafting implementation plan via writing-plans skill.

Implemented tokenizer pipeline split with new stage modules: `src/core/services/tokenizer/parser-selection-stage.ts`, `src/core/services/tokenizer/parser-enrichment-stage.ts`, `src/core/services/tokenizer/annotation-stage.ts`, and parser lifecycle runtime in `src/core/services/tokenizer/yomitan-parser-runtime.ts`; reduced `src/core/services/tokenizer.ts` to orchestration facade over stages.

Added direct stage-level tests: `src/core/services/tokenizer/parser-selection-stage.test.ts`, `src/core/services/tokenizer/parser-enrichment-stage.test.ts`, and `src/core/services/tokenizer/annotation-stage.test.ts`, and wired them into `test:core:src` + `test:core:dist` scripts in `package.json`.

Verification: `bun test src/core/services/tokenizer.test.ts src/core/services/tokenizer/annotation-stage.test.ts src/core/services/tokenizer/parser-selection-stage.test.ts src/core/services/tokenizer/parser-enrichment-stage.test.ts` PASS (53/53); `bun run test:core:src` PASS (219 pass, 6 skip); `bun run build` PASS; `bun run test:core:dist` PASS (214 pass, 10 skip).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Split tokenizer internals into explicit stages while preserving external behavior: parser candidate mapping/selection moved to `parser-selection-stage`, MeCab POS1 enrichment moved to `parser-enrichment-stage`, post-token annotation (known-word, frequency, JLPT, N+1) moved to `annotation-stage`, and Yomitan parser window lifecycle isolated in `yomitan-parser-runtime`. `src/core/services/tokenizer.ts` now acts as a thin orchestrator that normalizes subtitle text, requests parser output, runs stage pipeline, and handles MeCab fallback.

Added direct stage-level tests for scoring/selection and annotation semantics (`parser-selection-stage.test.ts`, `parser-enrichment-stage.test.ts`, `annotation-stage.test.ts`) and included them in both source and dist core test lanes via `package.json`. Validation passed across targeted tokenizer tests plus full core gates (`test:core:src`, `build`, `test:core:dist`) with no tokenizer regression.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Tokenizer-related test suites pass
- [x] #2 New stage-level tests exist for scoring and annotation
<!-- DOD:END -->
