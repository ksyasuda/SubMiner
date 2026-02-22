---
id: TASK-105
title: Eliminate unsafe non-test runtime casts in main boundaries
status: To Do
assignee: []
created_date: '2026-02-22 07:13'
updated_date: '2026-02-22 07:13'
labels:
  - refactor
  - type-safety
  - maintainability
dependencies:
  - TASK-97
  - TASK-80
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Production runtime wiring still relies on many unsafe type escapes (`as never`, `as unknown as`) across `src/main.ts` and `src/main/runtime/*` non-test files.

Current scan shows repeated casts in dependency builders and runtime adapters, which hides contract drift and allowed current Jellyfin compile break to persist until full build.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Build a baseline list of non-test unsafe casts in `src/main.ts` and `src/main/runtime`.
2. Group casts by ownership domain (Jellyfin, overlay, startup, IPC, tokenization, tray).
3. For each domain, introduce explicit interface types/adapters so casts are removed at boundary creation.
4. Keep test-only casts allowed where practical, but remove production-path `as never` usage.
5. Add compile-time contract assertions for critical dependency builders to catch drift early.
6. Validate with `bun run build` and affected source test suites.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Non-test `as never` occurrences in `src/main.ts` and `src/main/runtime` are reduced to zero or documented narrow exceptions.
- [ ] #2 Runtime dependency builders compile without unsafe production-path cast escapes.
- [ ] #3 Contract regressions are caught by compile/test checks rather than runtime behavior.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Cast reduction report attached in task notes (before/after counts).
- [ ] #2 `bun run build` and `bun run test:core:src` pass.
- [ ] #3 Any remaining exceptions have explicit rationale in code comments or task notes.
<!-- DOD:END -->

