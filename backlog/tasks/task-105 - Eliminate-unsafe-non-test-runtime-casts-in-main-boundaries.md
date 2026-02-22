---
id: TASK-105
title: Eliminate unsafe non-test runtime casts in main boundaries
status: Done
assignee:
  - opencode-task105-unsafe-casts
created_date: '2026-02-22 07:13'
updated_date: '2026-02-22 21:56'
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
- [x] #1 Non-test `as never` occurrences in `src/main.ts` and `src/main/runtime` are reduced to zero or documented narrow exceptions.
- [x] #2 Runtime dependency builders compile without unsafe production-path cast escapes.
- [x] #3 Contract regressions are caught by compile/test checks rather than runtime behavior.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1) Baseline cast inventory in non-test scope (`src/main.ts`, `src/main/runtime/*`) using `rg` and record before-count in notes.
2) Tighten runtime adapter contracts to remove `as never` escapes in `*-main-deps` modules: MPV/Jellyfin defaults, startup config deps, overlay visibility/options deps, field-grouping deps, CLI context deps, subtitle tokenization deps, dictionary runtime deps, tray/app deps, and MPV event bindings.
3) Update `src/main.ts` boundary wiring to match tightened contracts and remove cast escapes in config hot-reload setters, dictionary lookup setters, MPV defaults/OSD wiring, MPV client factory constructor typing, tray creation, and overlay bootstrap callbacks.
4) Re-run cast scan expecting zero non-test unsafe casts (or document narrow exceptions), run required gates (`bun run build`, `bun run test:core:src`), then attach before/after cast report and finalize AC/DoD.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Executed TASK-105 implementation with runtime contract tightening across `src/main.ts` + `src/main/runtime/*` non-test files. Removed all unsafe `as never` / `as unknown as` casts in target scope by replacing unknown passthrough contracts with typed deps (MPV/Jellyfin runtime types, overlay/bootstrap deps, tokenizer/CLI deps, dictionary lookup function types, tray deps typing).

Cast reduction report: baseline scan count = 42 non-test matches (`rg -n "as\\s+never|as\\s+unknown\\s+as" src/main.ts src/main/runtime --glob '!**/*.test.ts'`), after changes = 0 matches with same command.

Verification: targeted runtime cast-refactor tests passed (`bun test ...` across 19 files, 35 pass / 0 fail). Required core lane passed: `bun run test:core:src` => 236 pass / 6 skip / 0 fail.

Blocker: `bun run build` currently fails on pre-existing unrelated typing issues in `src/anki-integration/note-update-workflow.test.ts` (6 TS errors). No remaining build errors from TASK-105 touched files.

Follow-up verification rerun after concurrent compile-fix pass: `bun run build` now succeeds on current HEAD (tsc + renderer build + packaging copy step). DoD build gate unblocked.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Eliminated unsafe non-test runtime cast escapes in `src/main.ts` and `src/main/runtime/*` by tightening runtime dependency-builder contracts (MPV/Jellyfin, overlay/bootstrap, CLI/tokenizer, dictionary/tray boundaries) and removing `as never` / `as unknown as` patterns in production paths. Captured cast reduction evidence (42 -> 0 matches in scoped scan), validated focused runtime suites (35 pass), and revalidated required gates (`bun run test:core:src`, `bun run build`) on current HEAD.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Cast reduction report attached in task notes (before/after counts).
- [x] #2 `bun run build` and `bun run test:core:src` pass.
- [x] #3 Any remaining exceptions have explicit rationale in code comments or task notes.
<!-- DOD:END -->
