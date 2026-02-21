# codex-duplicate-kiku-20260221T043006Z-5vkz

- alias: `codex-duplicate-kiku`
- mission: `Fix Kiku duplicate-card detection/grouping regression for Yomitan duplicate-marked + N+1-highlighted cards`
- status: `completed`
- start_utc: `2026-02-21T04:30:06Z`
- last_update_utc: `2026-02-21T10:07:58Z`

## Intent

- Reproduce bug where clear duplicate cards no longer trigger Kiku duplicate grouping.
- Add failing regression test first (TDD).
- Patch duplicate detection logic with minimal behavior change.

## Planned Files

- `src/anki-integration/duplicate.ts`
- `src/anki-integration/duplicate.test.ts` (or nearest duplicate-detection tests)
- `docs/subagents/INDEX.md`
- `docs/subagents/collaboration.md`
- `backlog/tasks/task-94 - Fix-Kiku-duplicate-detection-for-Yomitan-marked-duplicates.md`

## Assumptions

- Duplicate signal should still come from Anki duplicate search + Yomitan/N+1-derived fields used in note content.
- Regression likely from term/readings normalization/query escaping mismatch.

## Outcome

- Root cause: candidate-note exact-check only resolved the originating field name (`Expression` or `Word`), so duplicates failed when candidate note used the opposite alias.
- Added regression test first (RED): `Expression` current note vs `Word` candidate with same value returned `null`.
- Implemented minimal fix: candidate resolution now checks both aliases (`word` and `expression`) before exact-value compare.
- GREEN: targeted duplicate test passed; related `anki-integration` test passed.
- User follow-up repro showed remaining miss when duplicate appears only in alias field search results.
- Added second RED test for alias-query fallback.
- Implemented query-stage alias fallback: run `findNotes` for both alias fields, merge note ids, then exact-verify.
- GREEN after follow-up: duplicate tests + `anki-integration` test pass.
- User reported still failing after first follow-up.
- Added third RED regression: source note containing both `Expression` (sentence) and `Word` (term) only matched duplicates via `Word`; previous logic missed this by using only one source value.
- Implemented source-candidate expansion: gather both `Word` and `Expression` source values, query aliases for each, dedupe queries, then exact-match against normalized set.
- GREEN: duplicate tests (3/3) + `anki-integration` test pass.
- Image-backed repro indicated possible duplicate outside configured deck scope.
- Added fourth RED regression: deck-scoped query misses, collection-wide query should still detect duplicate.
- Implemented deck fallback query pass (same source/alias combinations without deck filter) when deck-scoped pass yields no candidates.
- GREEN: duplicate tests (4/4) + `anki-integration` test pass.
- User confirmed fresh build/install still failed with `貴様` repro.
- Added fifth RED regression: field-specific queries return no matches but plain text query returns candidate.
- Implemented plain-text query fallback pass (deck-scoped then global), still gated by exact `word`/`expression` value verify.
- GREEN: duplicate tests (5/5) + `anki-integration` test pass.
- Added runtime debug instrumentation for duplicate detection query/verification path:
  - query string + hit count
  - candidate count after exclude
  - exact-match note id + field
- No behavior change from instrumentation; build + tests still green.
- User requested logging policy update: prefer console output unless explicitly captured, and persistent logs under `~/.config/SubMiner/logs/*.log`.
- Updated default launcher/app mpv log path to daily file naming: `~/.config/SubMiner/logs/SubMiner-YYYY-MM-DD.log`.
- Typecheck green.
- Found observability gap: app logger wrote only to stdout/stderr while launcher log file only captured wrapper messages.
- Added file sink to `src/logger.ts` so app logs also append to `~/.config/SubMiner/logs/SubMiner-YYYY-MM-DD.log` (or `SUBMINER_MPV_LOG` when set).
- Verified with typecheck + build.

## Files Touched

- `src/anki-integration/duplicate.ts`
- `src/anki-integration/duplicate.test.ts`
- `backlog/tasks/task-94 - Fix-Kiku-duplicate-detection-for-Yomitan-marked-duplicates.md`
- `docs/subagents/INDEX.md`
- `docs/subagents/collaboration.md`
- `docs/subagents/agents/codex-duplicate-kiku-20260221T043006Z-5vkz.md`

## Handoff

- No blockers.
- Next step: run broader gate (`bun run test:fast`) when ready, then commit.
