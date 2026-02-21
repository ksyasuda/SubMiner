---
id: TASK-94
title: Fix Kiku duplicate detection for Yomitan-marked duplicates
status: Done
assignee:
  - codex-duplicate-kiku-20260221T043006Z-5vkz
created_date: '2026-02-21 04:33'
updated_date: '2026-02-21 01:40'
labels:
  - bug
  - anki
  - kiku
dependencies: []
priority: high
ordinal: 65000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Kiku field grouping no longer detects duplicate cards in scenarios where the mined card is clearly marked duplicate by Yomitan/N+1 workflow. Restore duplicate detection so duplicate note lookup succeeds for equivalent expression/word cards and Kiku grouping can run.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Repro case covered by automated regression test in duplicate-detection path.
- [x] #2 Kiku duplicate detection returns duplicate note id for the repro case.
- [x] #3 Targeted tests for duplicate detection pass.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added regression test `src/anki-integration/duplicate.test.ts` for a cross-field duplicate case where current note uses `Expression` and candidate uses `Word` with same value.

Updated duplicate matching in `src/anki-integration/duplicate.ts` to try alternate field-name aliases (`word` <-> `expression`) when resolving candidate note fields for exact-value verification.

Follow-up fix: duplicate search query now also probes alias fields (`word` <-> `expression`) and merges candidate note ids before exact verification, so duplicates are still found when only the alias field is indexed/populated on existing cards.

Second follow-up fix: duplicate detection now evaluates both source values when current note contains both `Expression` and `Word` (previously only one was used, depending on field-order). Query and exact verification now run against all source duplicate candidates.

Third follow-up fix: if deck-scoped duplicate queries return no results, detection now retries the same source/alias query set collection-wide (no deck filter) before exact verification. This aligns with cases where Yomitan shows duplicates outside the configured mining deck.

Fourth follow-up fix: if field-specific queries miss entirely, detection now falls back to phrase/plain-text queries (deck-scoped then collection-wide) and still requires exact `Expression/Word` value verification before selecting a duplicate note.

Fifth follow-up: added explicit duplicate-search debug logs (query strings, hit counts, candidate counts, exact-match note id) to improve runtime diagnosis in live launcher runs.

Verification:
- `bun run build`
- `node dist/anki-integration/duplicate.test.js`
- `node --test dist/anki-integration.test.js`
<!-- SECTION:NOTES:END -->
