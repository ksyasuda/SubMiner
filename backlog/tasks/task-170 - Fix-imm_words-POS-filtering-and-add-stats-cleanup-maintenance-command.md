---
id: TASK-170
title: Fix imm_words POS filtering and add stats cleanup maintenance command
status: Done
assignee: []
created_date: '2026-03-13 00:00'
updated_date: '2026-03-18 05:31'
labels: []
milestone: m-1
dependencies: []
priority: high
ordinal: 9010
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`imm_words` is currently populated from raw subtitle text instead of tokenized subtitle metadata, so ignored functional/noise tokens leak into stats and no POS metadata is stored. Fix live persistence to follow the existing token annotation exclusion rules and add an on-demand stats cleanup command to remove stale bad vocabulary rows from the stats DB.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 New `imm_words` inserts use tokenized subtitle data, persist POS metadata, and skip tokens excluded by existing POS-based vocabulary ignore rules.
- [x] #2 `subminer stats cleanup` supports `-v` / `--vocab`, defaults to vocab cleanup, and removes stale bad `imm_words` rows on demand.
- [x] #3 Regression coverage exists for persistence filtering, cleanup behavior, and stats cleanup CLI wiring.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed `imm_words` persistence so the tracker now consumes tokenized subtitle data, stores POS metadata (`part_of_speech`, `pos1`, `pos2`, `pos3`), preserves distinct surface/lemma fields (`word` vs `headword`) when tokenization provides them, and skips vocabulary rows excluded by the existing POS/noise rules instead of mining raw subtitle fragments. Added `subminer stats cleanup` with default vocab cleanup plus `-v/--vocab`; the cleanup pass now repairs stale `headword`, `reading`, and `part_of_speech` values, attempts best-effort MeCab backfill for legacy rows, and removes rows that still have no usable POS metadata or fail the vocab filters.

Verification:

- `bun run typecheck`
- `bun test src/core/services/immersion-tracker-service.test.ts src/core/services/immersion-tracker/__tests__/query.test.ts src/core/services/immersion-tracker/storage-session.test.ts launcher/parse-args.test.ts launcher/commands/command-modules.test.ts src/main/runtime/stats-cli-command.test.ts src/main/runtime/mpv-main-event-main-deps.test.ts src/core/services/cli-command.test.ts`
- `bun run docs:test`
- `bun run docs:build`
<!-- SECTION:FINAL_SUMMARY:END -->
