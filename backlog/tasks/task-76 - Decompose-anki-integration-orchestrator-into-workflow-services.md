---
id: TASK-76
title: Decompose anki-integration orchestrator into workflow services
status: To Do
assignee: []
created_date: '2026-02-18 11:43'
updated_date: '2026-02-18 11:43'
labels:
  - anki
  - refactor
  - maintainability
dependencies:
  - TASK-57
priority: high
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
- [ ] #1 `src/anki-integration.ts` reduced to orchestration/composition role
- [ ] #2 Extracted workflow modules have focused tests
- [ ] #3 Existing mining behavior remains unchanged in regression tests
- [ ] #4 Documentation updated with ownership boundaries
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 `bun run test:core:dist` passes with Anki-related suites green
- [ ] #2 No callsite API breakage outside planned changes
<!-- DOD:END -->

