---
id: TASK-98
title: Gate subtitle character-name highlighting on character dictionary enablement
status: Done
assignee:
  - codex
created_date: '2026-03-07 00:54'
updated_date: '2026-03-16 05:13'
labels:
  - subtitle
  - character-dictionary
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/src/main.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/core/services/tokenizer.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/config/definitions/defaults-subtitle.ts
priority: medium
ordinal: 74500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Ensure subtitle tokenization and other annotations continue to work, but character-name lookup/highlighting is disabled whenever the AniList character dictionary feature is disabled. This avoids unnecessary name-match processing when the backing dictionary is unavailable.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 When anilist.characterDictionary.enabled is false, subtitle tokenization does not request character-name match metadata or highlight character names.
- [x] #2 When anilist.characterDictionary.enabled is true and subtitleStyle.nameMatchEnabled is true, existing character-name matching behavior remains enabled.
- [x] #3 Subtitle tokenization, JLPT, frequency, and other non-name annotation behavior remain unchanged when character dictionaries are disabled.
- [x] #4 Automated tests cover the runtime gating behavior.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a failing test in `src/main/runtime/subtitle-tokenization-main-deps.test.ts` proving name-match enablement resolves to false when `anilist.characterDictionary.enabled` is false even if `subtitleStyle.nameMatchEnabled` is true.
2. Update `src/main/runtime/subtitle-tokenization-main-deps.ts` and `src/main.ts` so subtitle tokenization only enables name matching when both the subtitle setting and the character dictionary setting are enabled.
3. Run focused Bun tests for the updated runtime deps and subtitle processing seams.
4. If verification stays green, check off acceptance criteria and record the result.

Implementation plan saved in `docs/plans/2026-03-06-character-name-gating.md`.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Created plan doc `docs/plans/2026-03-06-character-name-gating.md` after user approved the narrow runtime-gating approach. Proceeding with TDD from the subtitle tokenization main-deps seam.

Implemented the gate at the subtitle tokenization runtime-deps boundary so `getNameMatchEnabled` is false unless both `subtitleStyle.nameMatchEnabled` and `anilist.characterDictionary.enabled` are true.

Verification: `bun test src/main/runtime/subtitle-tokenization-main-deps.test.ts`, `bun test src/core/services/subtitle-processing-controller.test.ts`, `bun run typecheck`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Character-name lookup/highlighting is now suppressed when the AniList character dictionary is disabled, while subtitle tokenization and other annotation paths remain active. Added focused runtime-deps coverage and wired the main runtime to pass the character-dictionary enabled flag into subtitle tokenization.
<!-- SECTION:FINAL_SUMMARY:END -->
