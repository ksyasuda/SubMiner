---
id: TASK-350
title: Fix known highlighting for Yomitan compound tokens
status: Done
assignee:
  - codex
created_date: '2026-05-12 09:08'
updated_date: '2026-05-12 09:29'
labels:
  - bug
  - tokenizer
dependencies: []
modified_files:
  - src/core/services/tokenizer/yomitan-parser-runtime.ts
  - src/core/services/tokenizer/yomitan-parser-runtime.test.ts
  - src/core/services/tokenizer.test.ts
  - changes/350-known-yomitan-token-highlighting.md
priority: high
ordinal: 184500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Subtitle known-word coloring should respect the lexical token selected by Yomitan. If Yomitan emits a compound or inflected expression as one token, SubMiner must not mark that displayed token known solely because MeCab/POS enrichment can decompose it into known component words.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 A Yomitan token such as `取り組んで` with headword `取り組む` remains not-known when only component words like `取る` or `組む` are known.
- [x] #2 Frequency/JLPT/POS enrichment still works for the selected Yomitan token without leaking component known-word status into `isKnown`.
- [x] #3 Regression coverage demonstrates the compound-token case and fails on current behavior before the fix.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a regression in `src/core/services/tokenizer.test.ts` for a Yomitan-selected compound token: Yomitan emits `取り組んで` with headword `取り組む`; MeCab splits the same span into component tokens whose headwords include known component words such as `組む`; expected result is one displayed token with `isKnown === false` when only the components are known.
2. Verify the regression fails on current code.
3. Patch MeCab enrichment so it only contributes POS metadata used by annotation filters/exclusions. It must preserve the Yomitan token's `surface`, `headword`, `reading`, offsets, and existing lexical annotation state, especially `isKnown`.
4. Re-run the targeted tokenizer test, then a relevant fast test lane if practical.

After inspecting code, MeCab enrichment currently only writes POS metadata. The observed component coloring can also come from SubMiner's custom Yomitan scanning path fragmenting a phrase differently than Yomitan's internal parser. Regression should exercise `requestYomitanScanTokens` fallback/parser behavior as seen by `tokenizeSubtitle`, and the fix should prefer Yomitan internal parse token identity while keeping MeCab limited to filtering/POS metadata.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
User clarified MeCab is intended only to help filter unwanted characters/particles/sound effects/etc., not to alter lexical tokenization or known-word decisions.

Implementation settled on parse-first token identity: `requestYomitanScanTokens` now reads Yomitan internal parse tokens first. It still runs the scanner to keep scanner metadata when spans agree, but returns parse tokens when the scanner fragments the parse token. MeCab remains POS/filter enrichment only.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed known-word highlighting for Yomitan compound tokens by preferring Yomitan internal parse token spans over fragmented scanner output. When scanner output agrees with parse spans, scanner metadata such as name-match and word classes is preserved; when it fragments a Yomitan token, the parse token identity wins so known component words do not color the larger unknown token green.

Added regressions for `取り組んで` with known component words (`取る`, `組む`, `もらう`) and for parser-runtime token selection/metadata behavior. Added a changelog fragment.

Validation run: `bun test src/core/services/tokenizer.test.ts src/core/services/tokenizer/yomitan-parser-runtime.test.ts src/core/services/tokenizer/parser-selection-stage.test.ts src/core/services/tokenizer/parser-enrichment-stage.test.ts`; `bun run typecheck`; `bun x prettier --check src/core/services/tokenizer.test.ts src/core/services/tokenizer/yomitan-parser-runtime.ts src/core/services/tokenizer/yomitan-parser-runtime.test.ts changes/350-known-yomitan-token-highlighting.md`; `bun run changelog:lint`; `git diff --check`.
<!-- SECTION:FINAL_SUMMARY:END -->
