---
id: TASK-238.4
title: 'Decompose character dictionary runtime into fetch, build, and cache modules'
status: Done
assignee: []
created_date: '2026-03-26 20:49'
updated_date: '2026-03-31 19:37'
labels:
  - tech-debt
  - runtime
  - anilist
  - maintainability
milestone: m-0
dependencies:
  - TASK-238.3
references:
  - src/main/character-dictionary-runtime.ts
  - src/main/runtime/character-dictionary-auto-sync.ts
  - docs/architecture/README.md
parent_task_id: TASK-238
priority: medium
ordinal: 173500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`src/main/character-dictionary-runtime.ts` is now one of the largest live production files in the repo and combines AniList transport, name normalization, snapshot/image shaping, cache management, and zip packaging. That file will keep growing as character-dictionary features evolve. Split it into focused modules so the runtime surface becomes orchestration instead of a catch-all implementation blob.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 AniList fetch/parsing logic, dictionary-entry building, and snapshot/cache/zip persistence no longer live in one giant file.
- [x] #2 The public runtime API stays behavior-compatible for current callers.
- [x] #3 The top-level runtime/orchestration file becomes materially smaller and easier to review.
- [x] #4 Existing character-dictionary tests still pass, and new focused tests cover the extracted modules where needed.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Identify the dominant concern boundaries inside `src/main/character-dictionary-runtime.ts`.
2. Extract fetch/transform/persist modules with narrow interfaces, keeping data-shape ownership explicit.
3. Leave the exported runtime API stable for current main-process callers.
4. Verify with the maintained character-dictionary/runtime test lane plus `bun run typecheck`.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Split `src/main/character-dictionary-runtime.ts` into focused modules under `src/main/character-dictionary-runtime/` (`fetch`, `build`, `cache`, plus helper modules). The orchestrator stayed as a compatibility shim/API surface with delegated module functions. Added focused tests for cache snapshot semantics and term rebuild + collapsible-open-state behavior in the new modules. Updated runtime architecture docs in `docs/architecture/domains.md` and `docs-site/architecture.md`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Task completed with commit `5b06579e` (`refactor: split character dictionary runtime modules`). Runtime refactor landed with regression coverage and verification including runtime-compat lanes, and all changed behavior was validated as API-compatible for callers.
<!-- SECTION:FINAL_SUMMARY:END -->
