---
id: TASK-82
title: 'Subtitle frequency highlighting: fix noisy Yomitan readings and restore known/N+1 color priority'
status: Done
assignee: []
created_date: '2026-03-02 20:10'
updated_date: '2026-03-02 01:44'
labels: []
dependencies: []
priority: high
ordinal: 9002
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Address frequency-highlighting regressions:

- tokens like `断じて` missed rank assignment when Yomitan merged-token reading was truncated/noisy;
- known/N+1 tokens were incorrectly colored by frequency color instead of known/N+1 color.

Expected behavior:

- known/N+1 color always wins;
- if token is frequent and within `topX`, frequency rank label can still appear on hover/metadata.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Frequency lookup succeeds for noisy/truncated merged-token readings via robust fallback behavior.
- [x] #2 Merged-token reading normalization restores missing kana suffixes where safe (`headword === surface` path).
- [x] #3 Known/N+1 tokens keep known/N+1 color classes; frequency color class does not override them.
- [x] #4 Frequency rank hover label remains available for in-range frequent tokens, including known/N+1.
- [x] #5 Regression tests added for tokenizer and renderer behavior.

<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Implemented and validated:

- tokenizer now normalizes selected Yomitan merged-token readings by appending missing trailing kana suffixes when safe (`headword === surface`);
- frequency lookup now does lazy fallback: requests `{term, reading}` first, and only requests `{term, reading: null}` for misses;
- this removes eager `(term, null)` payload inflation on medium-frequency lines and reduces extension RPC payload/load;
- renderer restored known/N+1 color priority over frequency class coloring;
- frequency rank label display remains available for frequent known/N+1 tokens;
- added regression tests covering noisy-reading fallback, lazy fallback-query behavior, and renderer class/label precedence.

Related commits:

- `17a417e` (`fix(subtitle): improve frequency highlight reliability`)
- `79f37f3` (`fix(subtitle): prioritize known and n+1 colors over frequency`)

<!-- SECTION:FINAL_SUMMARY:END -->
