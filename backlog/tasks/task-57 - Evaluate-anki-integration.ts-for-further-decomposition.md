---
id: TASK-57
title: Evaluate anki-integration.ts for further decomposition
status: To Do
assignee: []
created_date: '2026-02-16 04:47'
labels: []
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/anki-integration.ts
  - /home/sudacode/projects/japanese/SubMiner/src/anki-integration/
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

anki-integration.ts remains the largest file at 1930 lines despite having domain modules extracted to `src/anki-integration/` directory.

Current state:

- Main file: `src/anki-integration.ts` (1930 lines)
- Extracted modules:
  - card-creation.ts (727 lines)
  - known-word-cache.ts (398 lines)
  - field-grouping.ts (79 lines)
  - polling.ts (36 lines)
  - duplicate.ts (30 lines)
  - ui-feedback.ts (29 lines)
  - ai.ts (4.2k)

This task is to evaluate whether the main anki-integration.ts file can be further decomposed or if 1930 lines is acceptable for a main integration class.

Evaluation criteria:

1. Are there remaining cohesive units that could be extracted?
2. Is the remaining code primarily orchestration logic (which is acceptable to be longer)?
3. Would further splitting improve or hurt readability?
4. Are there internal classes or helpers that could be standalone?

Deliverable: A decision document with recommendations - either proceed with further decomposition or document why the current state is acceptable.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Review anki-integration.ts structure and identify remaining cohesive units
- [ ] #2 Evaluate if extraction would improve or hurt maintainability
- [ ] #3 Document decision with rationale
- [ ] #4 If proceeding: create extraction plan with specific modules to create
- [ ] #5 If not proceeding: document architectural justification for keeping as-is
<!-- AC:END -->
