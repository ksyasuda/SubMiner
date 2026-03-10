---
id: TASK-95
title: Invalidate old character dictionary snapshots after kana alias schema change
status: Done
assignee: []
created_date: '2026-03-06 09:25'
updated_date: '2026-03-06 09:28'
labels:
  - dictionary
  - cache
dependencies: []
references:
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/character-dictionary-runtime.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/character-dictionary-runtime.test.ts
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Bump character dictionary snapshot format/version so cached AniList snapshots created before kana alias generation are rebuilt automatically on next auto-sync or generation run.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Old cached character dictionary snapshots are treated as invalid after the schema/version bump
- [x] #2 Current snapshot generation tests cover rebuild behavior across version mismatch
- [x] #3 No manual cache deletion is required for users to pick up kana alias term generation
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Bumped the character dictionary snapshot format version so cached AniList snapshots created before kana alias generation are automatically treated as stale and rebuilt. Added regression coverage that seeds an older-format snapshot and verifies `getOrCreateCurrentSnapshot` fetches fresh data and overwrites the stale cache.

Verified with `bun test src/main/character-dictionary-runtime.test.ts` and `bun run tsc --noEmit`.

<!-- SECTION:FINAL_SUMMARY:END -->
