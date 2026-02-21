---
id: TASK-97
title: Normalize runtime composer contracts
status: Done
assignee:
  - opencode
created_date: '2026-02-21 07:15'
updated_date: '2026-02-21 10:07'
labels:
  - architecture
  - type-safety
  - maintainability
dependencies:
  - TASK-94
  - TASK-71
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Runtime composer interfaces currently allow optional/null drift. Standardize contracts via shared context types and stricter per-composer dependency interfaces.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Inventory all composer entrypoints in `src/main/runtime/composers/` and classify current optional/nullable fields.
2. Introduce shared `RuntimeContext` types in `src/main/runtime/` with explicit non-null guarantees where required.
3. Narrow each composer input type to only required fields; remove permissive `any`/optional drift.
4. Add compile-time contract tests (type assertions) and runtime seam tests for wiring regressions.
5. Update composition root adapters in `src/main.ts` and domain registries to satisfy strict contracts.
6. Run verification gate: `bun run build`, `bun run check:main-fanin`, `bun run test:core:dist`.
7. Document contract rules in architecture/development docs.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Shared runtime context/types exist and are consumed by all composers.
- [x] #2 Composer inputs reject missing required dependencies at compile time.
- [x] #3 No behavior regression in runtime wiring tests.
- [x] #4 Main fan-in guard remains within configured threshold.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add shared composer contract helpers (`ComposerInputs`, `ComposerOutputs`, `BuiltMainDeps`) in `src/main/runtime/composers/contracts.ts` and consume them across all runtime composers.
2. Tighten high-risk composer boundaries (`mpv-runtime-composer.ts`, `jellyfin-remote-composer.ts`, `ipc-runtime-composer.ts`) and align `src/main.ts` callsites with normalized contract types.
3. Add compile-time contract checks in `src/main/runtime/composers/composer-contracts.type-test.ts` for required composer deps.
4. Update composer contract conventions in docs (`docs/architecture.md`, `docs/development.md`) and validate gates (`bun run build`, `bun run check:main-fanin`, `bun run test:core:dist`).
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-02-21: started execution pass in current session; loaded task context and scanning composer contracts/tests before writing implementation plan.

Implemented shared runtime composer contract module and rewired all composer option/result types to use shared contract helpers.

Normalized MPV/Jellyfin/IPC composer boundaries, removed composer-level cast escape hatches where possible, and aligned `src/main.ts` runtime composer callsites to compile against stricter contracts.

Added compile-only type contract coverage in `src/main/runtime/composers/composer-contracts.type-test.ts` to assert required dependency keys fail when omitted.

Verification: `bun run build` PASS; focused composer tests PASS (`dist/main/runtime/composers/{mpv,jellyfin,ipc}-runtime-composer.test.js`); `bun run check:main-fanin` PASS (86 import lines, 10 runtime paths); `bun run test:core:dist` PASS (204 pass, 10 skipped).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Normalized runtime composer contracts by introducing shared contract helpers (`ComposerInputs`, `ComposerOutputs`, `BuiltMainDeps`) and applying them across all composer modules. Tightened MPV/Jellyfin/IPC composer interfaces and main callsites to enforce required dependency surfaces at compile time, added compile-only contract assertions, updated architecture/development conventions, and revalidated build/fan-in/core-runtime test gates with no behavior regressions.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Type-level contract tests added for composer boundaries.
- [x] #2 `bun run build`, `bun run check:main-fanin`, and `bun run test:core:dist` pass.
- [x] #3 Docs updated with composer contract conventions.
<!-- DOD:END -->
