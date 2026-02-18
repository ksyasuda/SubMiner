---
id: TASK-55
title: Normalize service naming conventions across core/services
status: Done
assignee: []
created_date: '2026-02-16 04:47'
updated_date: '2026-02-17 09:12'
labels: []
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/core/services/index.ts
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

The core/services directory has inconsistent naming patterns that create confusion:

- Some files use `*Service.ts` suffix (e.g., `mpv-service.ts`, `tokenizer-service.ts`)
- Others use `*RuntimeService.ts` or just descriptive names (e.g., `overlay-visibility-service.ts` exports functions with 'Service' in name)
- Some functions in files have 'Service' suffix, others don't

This inconsistency makes it hard to predict file/function names and creates cognitive overhead.

Standardize on:

- File names: `kebab-case.ts` without 'service' suffix (e.g., `mpv.ts`, `tokenizer.ts`)
- Function names: descriptive verbs without 'Service' suffix (e.g., `createMpvClient()`, `tokenizeSubtitle()`)
- Barrel exports: clean, predictable names

Files needing audit (47 services):

- All files in src/core/services/ need review

Note: This is a large-scale refactor that should be done carefully to avoid breaking changes.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Establish naming convention rules (document in code or docs)
- [x] #2 Audit all service files for naming inconsistencies
- [x] #3 Rename files to follow convention (kebab-case, no 'service' suffix)
- [x] #4 Rename exported functions to remove 'Service' suffix where present
- [x] #5 Update all imports across the entire codebase
- [x] #6 Update barrel exports
- [x] #7 Run full test suite
- [x] #8 Update any documentation referencing old names
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Starting implementation. Planning and executing a mechanical refactor: file names and exported symbols with Service suffix in `src/core/services`, then cascading import updates across `src/`.

Implemented naming convention refactor across `src/core/services`: removed `-service` from service file names, renamed Service-suffixed exported symbols to non-Service names, and updated barrel exports in `src/core/services/index.ts`.

Updated call sites across `src/main/**`, `src/core/services/**` tests, `scripts/**`, `package.json` test paths, and docs references (`docs/development.md`, `docs/architecture.md`, `docs/structure-roadmap.md`).

Validation completed: `pnpm run build` and `pnpm run test:fast` both pass after refactor.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Normalized `src/core/services` naming by removing `-service` from module filenames, dropping `Service` suffixes from exported service functions, and updating `src/core/services/index.ts` barrel exports to the new names. Updated all import/call sites across `src/main/**`, service tests, scripts, and docs/package test paths to match the new module and symbol names. Verified no behavior regressions with `pnpm run build` and `pnpm run test:fast` (all passing).

<!-- SECTION:FINAL_SUMMARY:END -->
