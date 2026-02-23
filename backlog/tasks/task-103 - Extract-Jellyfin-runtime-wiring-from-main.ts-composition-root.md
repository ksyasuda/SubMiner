---
id: TASK-103
title: Extract Jellyfin runtime wiring from main.ts composition root
status: Done
assignee:
  - codex-task103-jellyfin-main-composer-20260222T220441Z-m8p1
created_date: '2026-02-22 07:13'
updated_date: '2026-02-23 02:06'
labels:
  - refactor
  - maintainability
  - jellyfin
dependencies:
  - TASK-102
  - TASK-94
  - TASK-97
priority: medium
ordinal: 100000
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
- [x] #1 `src/main.ts` no longer contains the full Jellyfin deps-building block.
- [x] #2 New Jellyfin composer modules have focused tests covering handler wiring.
- [x] #3 `bun run check:main-fanin` stays green after extraction.
- [x] #4 `bun run build` and `bun run test:core:src` pass.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Plan artifact: `docs/plans/2026-02-22-task-103-jellyfin-runtime-wiring.md`.

Execution plan:
1) Add `composeJellyfinRuntimeHandlers` composer module + seam tests and composer barrel export.
2) Replace inline Jellyfin deps-building block in `src/main.ts` with one composer invocation while preserving existing callsites (`runJellyfinCommand`, `openJellyfinSetupWindow`, remote session lifecycle, playback).
3) Update architecture docs ownership bullets for Jellyfin composer boundary.
4) Run required gates: `bun run check:main-fanin`, `bun run build`, `bun run test:core:src`.
5) Record AC/DoD evidence and final summary in Backlog task (no commit).
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-02-22T22:04:41Z: Execution started via Backlog MCP workflow. Loading task context, creating written plan artifact, then implementing extraction and focused validations without commit.

2026-02-22T22:38:58Z: Added new Jellyfin composition entrypoint `src/main/runtime/composers/jellyfin-runtime-composer.ts` and seam test `src/main/runtime/composers/jellyfin-runtime-composer.test.ts`; exported via composer barrel.

2026-02-22T22:38:58Z: Replaced inline Jellyfin deps-building block in `src/main.ts` with a single `composeJellyfinRuntimeHandlers(...)` invocation returning config/client/playback/remote/session/command/setup handlers used by existing callsites.

2026-02-22T22:38:58Z: Validation: `bun test src/main/runtime/composers/jellyfin-runtime-composer.test.ts src/main/runtime/composers/jellyfin-remote-composer.test.ts` PASS (2/2), `bun run check:main-fanin` PASS (`15 import lines, 9 unique runtime paths`), `bun run test:core:src` PASS (241 pass, 6 skip).

2026-02-22T22:38:58Z: `bun run build` currently FAILS in current working tree due pre-existing duplicate/invalid `src/main.ts` imports unrelated to TASK-103 extraction (many TS2300/TS2724 import-surface errors already present in file).

2026-02-22T22:49:05Z: Re-ran required completion gates after user fix: `bun run build` PASS, `bun run test:core:src` PASS (241 pass, 6 skip), `bun run check:main-fanin` PASS (`14 import lines, 11 unique runtime paths`).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Extracted Jellyfin composition-root wiring from `src/main.ts` into a dedicated runtime composer entrypoint (`composeJellyfinRuntimeHandlers`) so main now delegates Jellyfin config resolution, client-info wiring, playback orchestration, remote session handlers, command dispatch, and setup-window wiring through a single composer surface. Added focused seam coverage in `src/main/runtime/composers/jellyfin-runtime-composer.test.ts`, retained existing `jellyfin-remote-composer` coverage, and updated architecture ownership docs to declare Jellyfin composer boundaries.

Validation run: `bun test src/main/runtime/composers/jellyfin-runtime-composer.test.ts src/main/runtime/composers/jellyfin-remote-composer.test.ts` (pass), `bun run check:main-fanin` (pass), `bun run build` (pass), and `bun run test:core:src` (pass).
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 `src/main.ts` LOC reduced materially from current baseline.
- [x] #2 Jellyfin runtime wiring is centralized in named composer module(s) with clear ownership docs.
- [x] #3 No behavior regressions in Jellyfin command/setup flows.
<!-- DOD:END -->
