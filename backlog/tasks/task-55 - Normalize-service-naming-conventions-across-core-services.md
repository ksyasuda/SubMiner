---
id: TASK-55
title: Normalize service naming conventions across core/services
status: To Do
assignee: []
created_date: '2026-02-16 04:47'
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
- [ ] #1 Establish naming convention rules (document in code or docs)
- [ ] #2 Audit all service files for naming inconsistencies
- [ ] #3 Rename files to follow convention (kebab-case, no 'service' suffix)
- [ ] #4 Rename exported functions to remove 'Service' suffix where present
- [ ] #5 Update all imports across the entire codebase
- [ ] #6 Update barrel exports
- [ ] #7 Run full test suite
- [ ] #8 Update any documentation referencing old names
<!-- AC:END -->
