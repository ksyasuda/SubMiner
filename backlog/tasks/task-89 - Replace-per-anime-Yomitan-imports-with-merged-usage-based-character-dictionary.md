---
id: TASK-89
title: Replace per-anime Yomitan imports with merged usage-based character dictionary
status: Done
assignee:
  - '@codex'
created_date: '2026-03-06 07:59'
updated_date: '2026-03-06 08:09'
labels:
  - character-dictionary
  - yomitan
  - anilist
dependencies: []
references:
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/character-dictionary-runtime.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/runtime/character-dictionary-auto-sync.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/config/definitions/defaults-integrations.ts
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Replace TTL-based per-anime character dictionary imports with a single merged Yomitan dictionary built from locally stored per-media metadata snapshots. Retain only most-recently-used anime up to configured maxLoaded, rebuild merged import when retained set membership/order changes, and avoid rebuilding on revisits that do not change the retained set.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Character dictionary retention becomes usage-based rather than TTL-based.
- [x] #2 Only one Yomitan character dictionary import is maintained and updated as a merged dictionary.
- [x] #3 Local storage keeps only metadata/snapshots needed to rebuild the merged dictionary; per-anime source zip cache is removed.
- [x] #4 Merged dictionary rebuild occurs when retained-set membership or order changes, not on unchanged revisits.
- [x] #5 Tests cover merged rebuild, MRU eviction, and no-op revisits.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Replaced per-media auto-sync imports with one merged Yomitan dictionary. Added snapshot persistence in `src/main/character-dictionary-runtime.ts` so auto-sync stores normalized per-media term/image metadata locally under `character-dictionaries/snapshots/` and rebuilds `merged.zip` from the MRU retained media ids.

Updated `src/main/runtime/character-dictionary-auto-sync.ts` to keep only MRU `activeMediaIds` plus merged revision/title state, rebuild/import the merged dictionary only when retained-set membership/order changes or the merged import is missing/stale, and skip rebuild on unchanged revisits.

Kept manual `generateForCurrentMedia` support by generating a one-off per-media zip from the stored snapshot, but removed the old per-media zip cache path from auto-sync state.

Updated config/help text to describe usage-based merged retention and mark legacy TTL/eviction knobs as ignored.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented MRU-based merged character dictionary sync. Auto-sync now stores per-media normalized snapshots locally, rebuilds a single merged Yomitan dictionary when the retained anime set/order changes, and keeps `maxLoaded` as the cap on most-recently-used anime included in that merged import. Unchanged revisits no longer rebuild/import the dictionary.

Validation: `bun test src/main/runtime/character-dictionary-auto-sync.test.ts src/main/character-dictionary-runtime.test.ts`, `bun run tsc --noEmit`.
<!-- SECTION:FINAL_SUMMARY:END -->
