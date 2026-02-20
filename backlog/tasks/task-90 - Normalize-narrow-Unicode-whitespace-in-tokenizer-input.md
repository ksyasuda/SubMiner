---
id: TASK-90
title: Normalize narrow Unicode whitespace in tokenizer input
status: Done
assignee: []
created_date: '2026-02-20 06:17'
updated_date: '2026-02-20 06:20'
labels: []
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Fix tokenizer behavior where subtitle lines containing narrow/invisible Unicode spacing between Japanese segments can be split/grouped incorrectly compared with normal space handling.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 A regression test reproduces the subtitle sample containing narrow/invisible Unicode spacing and fails before fix.
- [x] #2 Tokenizer normalization treats narrow/invisible spacing variants consistently with regular spacing for grouping/highlight behavior.
- [x] #3 Existing tokenizer tests still pass.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Linked from subagent session `codex-narrow-space-tokenizer-20260220T061716Z-p97s`.

Added `src/subtitle/stages/normalize.test.ts` regression for `\u200B` separator in subtitle sample and updated `normalizeTokenizerInput` to map `U+200B/U+2060/U+FEFF` to regular spaces before whitespace collapsing.

Validation:
- `bun run build && node --test dist/subtitle/stages/normalize.test.js`
- `node --test dist/core/services/tokenizer.test.js`
<!-- SECTION:NOTES:END -->
