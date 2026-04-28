---
id: TASK-310
title: Suppress N+1 highlight for kana-only candidate sentences
status: Done
assignee:
  - Codex
created_date: '2026-04-28 06:55'
updated_date: '2026-04-28 07:04'
labels:
  - tokenizer
  - n+1
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Reduce noisy N+1 subtitle annotations when the only unknown candidates in a sentence are kana-only hiragana or katakana words, such as mostly-kana subtitle lines where highlighting a particle/helper-like token is low value.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 N+1 annotation does not mark a kana-only unknown target when all N+1 candidates in the sentence are kana-only.
- [x] #2 N+1 annotation continues to mark kanji or mixed-script unknown targets in otherwise eligible sentences.
- [x] #3 A focused regression test covers the kana-only candidate case.
- [x] #4 N+1 minimum sentence word count excludes tokens stripped by the subtitle annotation filter, so filtered grammar/noise tokens cannot satisfy minSentenceWords.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Keep the existing N+1 target eligibility guard: kana-only subtitle surfaces do not become N+1 targets.
2. Add a focused regression in src/core/services/tokenizer/annotation-stage.test.ts proving annotation-filtered tokens do not count toward ankiConnect.nPlusOne.minSentenceWords.
3. Verify the new regression fails before code changes.
4. Patch src/token-merger.ts so the N+1 minimum sentence word count uses the same subtitle-annotation eligibility filter as annotation rendering, excluding filtered particles/auxiliaries/noise from the count.
5. Re-run focused tokenizer tests, then update TASK-310 acceptance criteria and final notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Initial context: current token-merger has an existing surface-level kana-only guard in isNPlusOneCandidateToken, added in commit 9e4ad907. Need decide whether to broaden behavior to lookup/headword forms or verify current behavior only.

Implemented by treating kana-only N+1 candidates as kana-only even when their token surface includes surrounding subtitle punctuation such as ellipsis or dashes. Focused regression was red before the token-merger change: スイッチ… was marked true, then passed after the guard update. test:env initially hit an unrelated immersion-tracker active_days timing/order failure and Bun follow-on loader error; the failing test passed in isolation and the full test:env rerun passed.

Reopened for follow-up scope: minSentenceWords must count annotation-eligible tokens only, not tokens stripped from annotation metadata.

Implemented follow-up minSentenceWords behavior: unknown tokens filtered from N+1 targeting no longer contribute to sentence length; known eligible tokens and true N+1 candidates still count.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Changed N+1 sentence-length counting so minSentenceWords only counts known eligible words and actual N+1 target candidates. Unknown tokens filtered from N+1 targeting, including kana-only unknowns, no longer pad a sentence into eligibility. Existing annotation-filtered particles/auxiliaries remain excluded. Added regression coverage for the filtered unknown padding case while preserving kanji/mixed-script target behavior.

Verification: new regression failed before implementation; `bun test src/core/services/tokenizer/annotation-stage.test.ts -t "N\\+1"` pass; full `bun test src/core/services/tokenizer/annotation-stage.test.ts` pass; `bun test src/core/services/tokenizer.test.ts -t "N\\+1"` pass; `bun run typecheck` pass.
<!-- SECTION:FINAL_SUMMARY:END -->
