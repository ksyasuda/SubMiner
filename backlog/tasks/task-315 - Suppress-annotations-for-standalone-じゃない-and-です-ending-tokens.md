---
id: TASK-315
title: Suppress annotations for standalone じゃない and です ending tokens
status: Done
assignee:
  - codex
created_date: '2026-05-03 00:02'
updated_date: '2026-05-03 00:31'
labels:
  - bug
  - tokenizer
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Standalone `じゃない` grammar ending tokens should not display or persist subtitle annotations even if a dictionary assigns a rank or JLPT/known match. User observed `じゃない` still being marked frequent in overlay after tokenization produced it as a dictionary word.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `じゃない` and `です` ending tokens have known-word, N+1, frequency, and JLPT annotation metadata cleared in subtitle annotation output.
- [x] #2 Common polite/question variants such as `じゃないですか` and `ですよ` remain excluded when tokenized as a single ending token.
- [x] #3 Regression coverage proves same-line Yomitan segments split content from trailing grammar endings so the content word can be annotated without coloring the ending.
- [x] #4 Auxiliary-only helper spans such as `てく` + `れた` in `ベアトリスがいてくれたから` have known-word, N+1, frequency, and JLPT annotation metadata cleared.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a focused regression for `ベアトリスがいてくれたから` where Yomitan tokens include auxiliary-only `てく` and `れた` with pre-ranked/known/JLPT metadata candidates.
2. Run the targeted test to verify the regression fails before production changes.
3. Patch the shared subtitle annotation filter so kana-only auxiliary helper spans made only of grammar POS components are excluded while preserving lexical content tokens.
4. Re-run targeted tokenizer/annotation tests, then run SubMiner change verification classifier/verifier for the touched files.
5. Update TASK-315 acceptance criteria, notes, and final summary with commands and outcomes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented as one focused tokenizer fix. Parser selection now splits dictionary-backed same-line grammar ending segments (`です`, `じゃない*`) from preceding content so annotation styling can apply only to the content token. Shared subtitle annotation filtering now treats bare `です` like the existing `ですか/ですよ/...` copula endings.

2026-05-03: Reopened for approved add-on covering auxiliary-only `てく` + `れた` helper highlighting report.

2026-05-03: Added regression coverage for `ベアトリスがいてくれたから` where Yomitan emits `てく` + `れた` and MeCab enrichment tags `てく` as `助詞|動詞` / `接続助詞|非自立`. The regression initially failed because `てく` kept `isKnown: true` and `jlptLevel: N4`. Added a shared-filter helper for kana-only particle+non-independent-verb helper spans, preserving lexical `自立` verbs. Verification: `bun test src/core/services/tokenizer/annotation-stage.test.ts`, `bun test src/core/services/tokenizer.test.ts`, `bun test src/core/services/tokenizer/parser-selection-stage.test.ts`, `bun x prettier --check ...`, and `bun run typecheck` passed. SubMiner verifier core lane passed typecheck but `bun run test:fast` failed on unrelated existing cross-suite issues: `window.electronAPI` undefined in `src/renderer/handlers/keyboard.ts` during `src/core/services/subsync.test.ts`, followed by Bun `node:test` nested-test cascade.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Split dictionary-backed trailing grammar ending segments (`です`, `じゃない*`) from preceding Yomitan same-line content before annotation, and added bare `です` to the explicit polite copula exclusion set.

Added the approved auxiliary-helper fix for `ベアトリスがいてくれたから`: kana-only `てく` + `れた` helper spans now clear known-word, N+1, frequency, and JLPT annotation metadata when POS enrichment shows a particle + non-independent verb helper, while lexical `自立` verb forms like `くれ`/`くれる` remain eligible.

Verification passed for targeted tokenizer/annotation/parser tests, Prettier check on touched files, and `bun run typecheck`. The SubMiner core verifier's `test:fast` step remains blocked by unrelated pre-existing cross-suite failures in `subsync`/renderer keyboard globals plus Bun `node:test` cascade; artifact: `.tmp/skill-verification/subminer-verify-20260502-173004-CMu3ai/`.
<!-- SECTION:FINAL_SUMMARY:END -->
