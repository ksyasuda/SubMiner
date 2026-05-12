---
id: TASK-312
title: Suppress ja-nai explanatory ending subtitle annotations
status: Done
assignee: []
created_date: '2026-05-02 09:55'
updated_date: '2026-05-02 10:03'
labels:
  - tokenizer
  - annotations
  - bug
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Suppress subtitle annotation styling for grammar-only explanatory endings like `じゃない` and `じゃないですか` while preserving nearby lexical content annotations.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `じゃない` and `じゃないですか`-style endings render as plain hoverable subtitle tokens.
- [x] #2 The reported phrase `みたいなのあるじゃないですか` does not annotate `じゃない`/`じゃないですか` as lexical/frequency content.
- [x] #3 Regression tests cover unit filter behavior and tokenizer integration without suppressing lexical content tokens.
- [x] #4 Standalone polite copula endings such as `です` / `ですよ` render as plain hoverable subtitle tokens even if POS metadata is missing or too lexical.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added failing coverage first for `じゃない` / `じゃないですか` and `ですよ` leaking annotation metadata when POS metadata is missing or too lexical. Implemented term-family exclusions in the shared subtitle annotation filter for the `じゃない` explanatory family and polite copula suffix endings (`ですか`, `ですね`, `ですよ`, `ですな`). Kept bare `です` term-only behavior unchanged to preserve existing no-POS frequency tests; POS-tagged `です` is already stripped by the grammar POS exclusion path.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Suppressed subtitle annotation metadata for grammar-only endings like `じゃないですか` and `ですよ`, while preserving nearby lexical content annotations. Added unit and tokenizer regression coverage for the reported `みたいなのあるじゃないですか` and `感じですよ` shapes, plus changelog fragment `changes/312-grammar-ending-annotation-filter.md`.

Validation: `bun test src/core/services/tokenizer/annotation-stage.test.ts`; `bun test src/core/services/tokenizer.test.ts`; `bun run typecheck`; `bun run changelog:lint`; `git diff --check`.
<!-- SECTION:FINAL_SUMMARY:END -->
