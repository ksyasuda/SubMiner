---
id: TASK-154
title: Avoid merged dictionary rebuilds on MRU reorder-only revisits
status: Done
assignee:
  - '@codex'
created_date: '2026-03-10 09:16'
updated_date: '2026-03-16 05:13'
labels:
  - character-dictionary
  - yomitan
  - anilist
dependencies: []
references:
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/runtime/character-dictionary-auto-sync.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/runtime/character-dictionary-auto-sync.test.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/character-dictionary-runtime.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/character-dictionary-runtime.test.ts
documentation:
  - /home/sudacode/projects/japanese/subminer-docs/development.md
  - /home/sudacode/projects/japanese/subminer-docs/architecture.md
priority: high
ordinal: 31500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Keep per-title MRU retention ordering for eviction decisions, but do not rebuild or reimport the merged character dictionary when the retained set membership is unchanged and only the MRU order changes.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Revisiting an anime already inside the retained set updates MRU ordering used for later eviction decisions without rebuilding or reimporting the merged dictionary when retained membership is unchanged.
- [x] #2 Merged dictionary revisions and contents are stable for the same retained membership regardless of MRU ordering.
- [x] #3 Adding a new anime that changes retained membership still rebuilds and imports the merged dictionary with the correct eviction behavior.
- [x] #4 Regression tests cover reorder-only revisits and stable merged revisions for equivalent retained sets.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a failing auto-sync regression showing that a revisit which only reorders MRU state should update `activeMediaIds` but skip merged rebuild/import.
2. Add a runtime-level regression showing `buildMergedDictionary` produces a stable revision for equivalent retained memberships regardless of input order.
3. Update merged-dictionary build normalization and auto-sync rebuild gating so rebuilds are driven by retained membership changes (or snapshot changes), not MRU reordering alone.
4. Re-run focused dictionary tests plus the local verification lanes impacted by the change.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Changed `src/main/runtime/character-dictionary-auto-sync.ts` so MRU order still updates in `activeMediaIds`, but merged rebuild/import now keys off retained membership changes rather than order-only reordering. Order-only revisits no longer force a merged ZIP rewrite or Yomitan reimport.

Changed `src/main/character-dictionary-runtime.ts` to canonicalize merged dictionary media ids before building, which makes merged revisions stable for equivalent retained memberships regardless of MRU ordering.

Updated regression coverage in `src/main/runtime/character-dictionary-auto-sync.test.ts` for reorder-only revisits and membership-change rebuilds, and extended `src/main/character-dictionary-runtime.test.ts` to assert stable merged revisions for reordered inputs. Verified with `bun test src/main/runtime/character-dictionary-auto-sync.test.ts src/main/character-dictionary-runtime.test.ts`, `bun run typecheck`, `bun run changelog:lint`, and `bun run test:fast`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Updated merged character-dictionary syncing so MRU ordering is still tracked for eviction in `src/main/runtime/character-dictionary-auto-sync.ts`, but reorder-only revisits no longer rebuild or reimport the merged dictionary unless retained membership or snapshot data changes. Canonicalized merged build input ordering in `src/main/character-dictionary-runtime.ts` so revisions remain stable for the same retained set regardless of MRU order, and added regressions in `src/main/runtime/character-dictionary-auto-sync.test.ts` plus `src/main/character-dictionary-runtime.test.ts`. Also updated `changes/character-dictionary-mru-state-recovery.md` to cover the new behavior.

Validation: `bun test src/main/runtime/character-dictionary-auto-sync.test.ts src/main/character-dictionary-runtime.test.ts`, `bun run typecheck`, `bun run changelog:lint`, `bun run test:fast`.
<!-- SECTION:FINAL_SUMMARY:END -->
