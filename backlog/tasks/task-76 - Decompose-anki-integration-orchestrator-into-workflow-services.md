---
id: TASK-76
title: Decompose anki-integration orchestrator into workflow services
status: Done
assignee: []
created_date: '2026-02-18 11:43'
updated_date: '2026-02-22 07:49'
labels:
  - anki
  - refactor
  - maintainability
dependencies:
  - TASK-57
priority: high
ordinal: 83000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`src/anki-integration.ts` remains a large orchestration hotspot after earlier splits. This task executes decomposition work (beyond evaluation) by extracting cohesive workflows and narrowing the central orchestrator to composition + policy decisions.
<!-- SECTION:DESCRIPTION:END -->

## Suggestions

<!-- SECTION:SUGGESTIONS:BEGIN -->
- Prefer workflow modules over helper sprawl: `card-update-workflow`, `media-attachment-workflow`, `known-word-sync-workflow`, `field-mapping-workflow`.
- Keep side-effect adapters explicit (AnkiConnect API, fs/media, notification).
- Introduce a small façade API from `anki-integration.ts` to preserve callsite stability.
<!-- SECTION:SUGGESTIONS:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Convert `TASK-57` findings into concrete extraction map with target module ownership.
2. Extract top two highest-churn workflows first (card creation/update and media attachment).
3. Move pure mapping/normalization logic to isolated pure modules with unit tests.
4. Keep backward-compatible public API for existing callsites.
5. Add regression tests for key mining flows (new card, update existing, audio-card mark, failure handling).
6. Update `docs/anki-integration.md` with new module boundaries.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `src/anki-integration.ts` reduced to orchestration/composition role
- [x] #2 Extracted workflow modules have focused tests
- [x] #3 Existing mining behavior remains unchanged in regression tests
- [x] #4 Documentation updated with ownership boundaries
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Execution started via writing-plans + executing-plans workflow.

Plan: docs/plans/2026-02-21-task-76-anki-workflow-services-plan.md

Baseline LOC (2026-02-21): src/anki-integration.ts=1315; existing anki-integration collaborators total=2271 (src/anki-integration/*.ts excluding facade).

Implemented workflow-service decomposition in `src/anki-integration.ts` by extracting `src/anki-integration/note-update-workflow.ts` (new-card update pipeline) and `src/anki-integration/field-grouping-workflow.ts` (auto/manual grouping merge orchestration).

Added focused workflow seam tests: `src/anki-integration/note-update-workflow.test.ts` and `src/anki-integration/field-grouping-workflow.test.ts`.

Updated ownership boundaries docs in `docs/anki-integration.md` under "Ownership Boundaries (TASK-76)".

Verification: `bun run build && node --test dist/anki-integration.test.js dist/anki-integration/note-update-workflow.test.js dist/anki-integration/field-grouping-workflow.test.js` (pass), and `bun run build && bun run test:core:dist` (pass).

LOC delta: `src/anki-integration.ts` 1315 -> 1143 (-172 LOC).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Decomposed `AnkiIntegration` into workflow services by extracting note update and field-grouping orchestration into dedicated modules while keeping public APIs stable. Added focused workflow seam tests, documented ownership boundaries, and validated behavior with Anki-focused dist tests plus full `test:core:dist` gate.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 `bun run test:core:dist` passes with Anki-related suites green
- [x] #2 No callsite API breakage outside planned changes
<!-- DOD:END -->
