---
id: TASK-332
title: Fix subtitle frequency annotation missing ranks shown in Yomitan popup
status: Done
assignee:
  - Codex
created_date: '2026-05-04 03:29'
updated_date: '2026-05-04 03:41'
labels:
  - bug
  - tokenizer
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Subtitle frequency highlighting can miss a token even when the Yomitan popup shows a rank within the configured threshold. Reproduced with `第二走者とアンカーは\n中継地点に速やかに移動！`: Yomitan popup shows `第二` JPDB rank 1820, but SubMiner tokenizer output has no `frequencyRank` for `第二`, so renderer cannot annotate it.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `第二` in `第二走者とアンカーは\n中継地点に速やかに移動！` receives the Yomitan rank shown by the popup when frequency highlighting is enabled.
- [x] #2 Regression test covers the Yomitan scan/frequency ingestion path for exact popup-derived ranks.
- [x] #3 Existing tokenizer frequency tests continue to pass.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Reproduce and inspect the missing `第二` rank path with tokenizer probes and focused tests.
2. Preserve exact Yomitan scan frequency ranks when the matching frequency entry omits reading metadata but has the same exact term.
3. Allow ranked ordinal prefix-noun tokens (`第` + numeric noun, e.g. `第二`) through annotation POS filtering while keeping standalone prefixes excluded.
4. Verify with focused tokenizer/runtime/annotation tests, typecheck, changelog lint, and a live-style Yomitan profile probe.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Root-cause probe against temp copy of Yomitan profile: tokenizer returns no frequencyRank for `第二`; renderer config `topX` is 10000, so render threshold is not the blocker.

User approved implementation plan on 2026-05-04.

Verification: `bun test src/core/services/tokenizer.test.ts src/core/services/tokenizer/yomitan-parser-runtime.test.ts src/core/services/tokenizer/annotation-stage.test.ts` passed (192 tests).

Verification: `bun run typecheck` passed.

Verification: `bun run changelog:lint` passed.

Verification: `bun run get-frequency:electron -- --yomitan-user-data /tmp/subminer-yomitan-probe-909423 "第二走者とアンカーは\\n中継地点に速やかに移動！"` produced `第二` with `frequencyRank: 1820`.

Finalization check: implementation plan updated to reflect the discovered POS-filter root cause and completed solution.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed subtitle frequency annotation for `第二` by allowing ranked ordinal prefix-noun compounds through annotation POS filtering. Also made scan rank matching tolerate exact frequency entries where one side omits reading metadata. Verified with tokenizer/runtime/annotation tests, typecheck, changelog lint, and a live-style Yomitan profile probe showing `第二` now receives frequencyRank 1820.
<!-- SECTION:FINAL_SUMMARY:END -->
