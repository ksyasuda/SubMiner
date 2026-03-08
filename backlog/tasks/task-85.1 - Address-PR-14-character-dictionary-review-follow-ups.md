---
id: TASK-85.1
title: 'Address PR #14 character dictionary review follow-ups'
status: Done
assignee:
  - codex
created_date: '2026-03-06 07:48'
updated_date: '2026-03-06 07:56'
labels: []
dependencies: []
references:
  - >-
    /home/sudacode/projects/japanese/SubMiner/launcher/commands/dictionary-command.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/character-dictionary-runtime.ts
  - /home/sudacode/projects/japanese/SubMiner/launcher/types.ts
documentation:
  - 'https://docs.anilist.co/guide/rate-limiting'
parent_task_id: TASK-85
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Apply the accepted follow-up fixes from Claude's PR review for the AniList character dictionary work: remove dead launcher code, deduplicate video extension handling where practical, and add explicit pacing for AniList character-page requests / character image downloads so the integration stays within AniList rate-limiting expectations.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launcher dictionary command no longer contains unreachable dead code after the app handoff.
- [x] #2 Character dictionary runtime no longer maintains a separate ad hoc video extension list when existing shared extension data can be reused safely.
- [x] #3 Character dictionary generation spaces outbound AniList-related requests with explicit named delays, and tests cover the pacing behavior and unchanged command forwarding behavior.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing tests for dictionary command handoff semantics and dictionary runtime request pacing.
2. Remove unreachable boolean return path from the launcher dictionary command while preserving call sites.
3. Reuse the shared launcher video extension set inside the character dictionary runtime with extname normalization, then add named AniList pacing constants for page fetches and character image downloads.
4. Run targeted tests, then broader relevant test slices, and update acceptance criteria / notes with the validated result.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added a shared `src/shared/video-extensions.ts` source and rewired both launcher/runtime consumers to remove the duplicated runtime extension list.

Replaced the hardcoded AniList page sleep with a per-generation AniList request pacer (2000ms between API requests) plus 250ms spacing between character image download attempts, including failed image fetches.

Hardened `runDictionaryCommand` so an unexpected return from the `never`-typed app handoff throws immediately instead of silently falling through.

Validated with targeted and adjacent test slices plus `bun run tsc --noEmit`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Removed the dead post-handoff return from the launcher dictionary command and replaced it with an explicit invariant error if the `never`-typed app handoff ever returns unexpectedly. Extracted video extension data into `src/shared/video-extensions.ts` so the launcher and character dictionary runtime share one source of truth.

Adjusted character dictionary generation to use a per-run AniList request pacer with a conservative 2000ms delay between AniList API calls, and added 250ms spacing between character image download attempts so repeated image fetches are not bursty even when an image URL fails. Added regression coverage for the pacing behavior and the launcher handoff invariant.

Validation: `bun test src/main/character-dictionary-runtime.test.ts`, `bun test launcher/commands/command-modules.test.ts`, `bun test launcher/main.test.ts launcher/parse-args.test.ts src/cli/args.test.ts src/core/services/cli-command.test.ts src/main/runtime/character-dictionary-auto-sync.test.ts`, `bun run tsc --noEmit`.
<!-- SECTION:FINAL_SUMMARY:END -->
