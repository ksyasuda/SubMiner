---
id: TASK-153
title: Fix character dictionary MRU eviction after revisits
status: Done
assignee:
  - '@codex'
created_date: '2026-03-10 07:56'
updated_date: '2026-03-10 08:48'
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
documentation:
  - /home/sudacode/projects/japanese/subminer-docs/development.md
  - /home/sudacode/projects/japanese/subminer-docs/architecture.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Keep merged character dictionary retention truly MRU-based when a currently retained anime is revisited before opening a new title, so the oldest retained title is evicted instead of the revisited one.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Revisiting an anime already retained in the merged character dictionary updates MRU ordering before the next new title is added.
- [x] #2 When retention exceeds maxLoaded after that revisit-plus-new-title sequence, the least recently used retained anime is evicted rather than the revisited title.
- [x] #3 Auto-sync rebuilds or reloads the merged dictionary so an evicted anime becomes available again when reopened.
- [x] #4 Regression tests cover the revisit-before-eviction flow.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Reproduce the revisit-before-eviction scenario with a focused regression test in `src/main/runtime/character-dictionary-auto-sync.test.ts` using a retained set that revisits an existing anime before introducing a new anime.
2. Run the focused test to confirm current behavior fails for the expected MRU ordering / eviction case.
3. Adjust `src/main/runtime/character-dictionary-auto-sync.ts` so retained ordering stays MRU-correct across revisits and subsequent additions without regressing unchanged revisit behavior.
4. Re-run the focused auto-sync suite, then run the relevant broader checks required for handoff.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added a focused regression in `src/main/runtime/character-dictionary-auto-sync.test.ts` for the `1,2,3,1,4,1` revisit-before-eviction flow. The current MRU auto-sync logic passed: merged builds were `[1]`, `[2,1]`, `[3,2,1]`, `[1,3,2]`, `[4,1,3]`, `[1,4,3]`, so I have not reproduced the reported eviction bug in the in-process auto-sync service yet.

Inspected local cache under `~/.config/SubMiner/character-dictionaries/` and found Little Witch Academia snapshot `anilist-21858.json` already present, so the reported regeneration path was not a snapshot-cache miss. The on-disk `merged.zip` revision (`9251ae23e136`) also differed from `auto-sync-state.json` (`bc16c5b5af17`), indicating a rebuilt merged dictionary artifact could land without the MRU state being persisted.

Fixed `src/main/runtime/character-dictionary-auto-sync.ts` to persist the rebuilt retained-set state immediately after a merged dictionary revision/title is known, before later Yomitan mutation steps. This keeps MRU eviction state aligned even if import/settings work fails after the rebuild artifact is written.

Added regression coverage in `src/main/runtime/character-dictionary-auto-sync.test.ts` for the revisit-before-eviction sequence and for post-build import failure preserving the rebuilt MRU state. Verified with `bun test src/main/runtime/character-dictionary-auto-sync.test.ts src/main/character-dictionary-runtime.test.ts`, `bun run typecheck`, `bun run changelog:lint`, and `bun run test:fast`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Persisted merged character-dictionary MRU state as soon as a rebuilt retained set is known in `src/main/runtime/character-dictionary-auto-sync.ts`, so revisits are not lost if later Yomitan import/settings work fails after `merged.zip` has already been rewritten. Added regression coverage for revisit-before-eviction ordering and import-failure state preservation in `src/main/runtime/character-dictionary-auto-sync.test.ts`, plus a changelog fragment in `changes/character-dictionary-mru-state-recovery.md`.

Validation: `bun test src/main/runtime/character-dictionary-auto-sync.test.ts src/main/character-dictionary-runtime.test.ts`, `bun run typecheck`, `bun run changelog:lint`, `bun run test:fast`.
<!-- SECTION:FINAL_SUMMARY:END -->
