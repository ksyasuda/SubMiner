---
id: TASK-103
title: Extract Jellyfin runtime wiring from main.ts composition root
status: To Do
assignee: []
created_date: '2026-02-22 07:13'
updated_date: '2026-02-22 07:13'
labels:
  - refactor
  - maintainability
  - jellyfin
dependencies:
  - TASK-102
  - TASK-94
  - TASK-97
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`src/main.ts` remains oversized (3014 LOC) and still carries a dense Jellyfin wiring block (roughly `main.ts:1145-1565`) despite composition-root refactors.

Goal: finish extraction of Jellyfin-specific dependency construction and command/setup orchestration into dedicated runtime composer modules so `main.ts` remains thin and legible.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Extract Jellyfin config/client-info/auth/list/play/setup dependency builders from `src/main.ts` into focused files under `src/main/runtime/composers/`.
2. Create one Jellyfin composition entrypoint that returns fully wired handlers used by `main.ts`.
3. Remove duplicated inline adapter lambdas from `main.ts`; keep only high-level invocation points.
4. Add/expand composer seam tests for extracted Jellyfin module boundaries.
5. Re-run fan-in and compile checks to ensure extraction does not regress runtime wiring.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 `src/main.ts` no longer contains the full Jellyfin deps-building block.
- [ ] #2 New Jellyfin composer modules have focused tests covering handler wiring.
- [ ] #3 `bun run check:main-fanin` stays green after extraction.
- [ ] #4 `bun run build` and `bun run test:core:src` pass.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 `src/main.ts` LOC reduced materially from current baseline.
- [ ] #2 Jellyfin runtime wiring is centralized in named composer module(s) with clear ownership docs.
- [ ] #3 No behavior regressions in Jellyfin command/setup flows.
<!-- DOD:END -->

