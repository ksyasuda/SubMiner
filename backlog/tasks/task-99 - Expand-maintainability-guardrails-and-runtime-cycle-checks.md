---
id: TASK-99
title: Expand maintainability guardrails and runtime cycle checks
status: Done
assignee: []
created_date: '2026-02-21 07:15'
updated_date: '2026-02-22 07:49'
labels:
  - quality
  - architecture
  - ci
dependencies:
  - TASK-96
  - TASK-97
priority: medium
ordinal: 71000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Current guardrails cover `main.ts` fan-in and broad file budgets. Add targeted checks for newly extracted hotspots and runtime import-cycle detection.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Rebaseline file-size budgets for known hotspots (including post-TASK-96 files).
2. Extend `check:file-budgets` config to enforce thresholds on new modules and key test files.
3. Add runtime dependency cycle detection for `src/main/runtime/**` (script + CI hook).
4. Wire new checks into local gate commands and CI workflow.
5. Add failure guidance output so contributors can remediate quickly.
6. Run gates on current branch and capture pass/fail evidence in task notes.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Budget thresholds include new hotspot modules and prevent silent growth.
- [x] #2 Runtime cycle check detects at least one injected fixture cycle in tests.
- [x] #3 CI runs both checks and fails fast on violations.
- [x] #4 Contributor guidance exists for fixing guardrail failures.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1) Extend `scripts/check-file-budgets.ts` with explicit hotspot thresholds (including `src/main.ts`, `src/config/service.ts`, `src/core/services/tokenizer.ts`, and `launcher/main.ts`), baseline-aware reporting, and strict failure on hotspot drift; update docs with baseline/threshold evidence.
2) Add `scripts/check-runtime-cycles.ts` to scan relative imports under `src/main/runtime/**`, detect cycles via DFS, and expose strict/non-strict modes; add fixture-backed tests in `scripts/check-runtime-cycles.test.ts` proving detection of an injected cycle.
3) Wire strict guardrails into CI fail-fast ordering (`check:file-budgets:strict`, `check:main-fanin:strict`, `check:runtime-cycles:strict`) before expensive test/build lanes.
4) Update contributor guidance in `docs/development.md` with local guardrail commands, expected outputs, and remediation flow.
5) Run verification (`check:file-budgets:strict`, `check:main-fanin:strict`, `check:runtime-cycles:strict`, `bun test scripts/check-runtime-cycles.test.ts`, `bun run test:fast`) and then finalize TASK-99 AC/DoD evidence in Backlog (no commit).
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-02-22T01:09:30Z: execution started in codex session codex-task99-guardrails-20260222T010930Z-m9q2; using writing-plans then executing-plans workflow with parallel subagents where safe.

2026-02-22T03:01:34Z: implemented hotspot budget guardrails in scripts/check-file-budgets.ts with explicit baseline/limit reporting for src/main.ts, src/config/service.ts, src/core/services/tokenizer.ts, launcher/main.ts, src/config/resolve.ts, and src/main/runtime/composers/contracts.ts.

2026-02-22T03:01:34Z: added scripts/check-runtime-cycles.ts + scripts/check-runtime-cycles.test.ts with fixture coverage under scripts/fixtures/runtime-cycles/{acyclic,cyclic}; strict mode exits non-zero on detected cycle and fixture command confirms detection path module-a.ts -> module-b.ts -> nested/index.ts -> module-a.ts.

2026-02-22T03:01:34Z: CI fail-fast guardrail step added in .github/workflows/ci.yml running check:file-budgets:strict, check:main-fanin:strict, and check:runtime-cycles:strict immediately after install.

2026-02-22T03:01:34Z: docs/development.md now includes Maintainability Guardrails section with local commands, expected [OK] output patterns, hotspot baseline/limit table, and remediation guidance.

Verification: bun run check:file-budgets:strict (passes hotspot gate; warns on legacy broad over-500 files), bun run check:main-fanin:strict (OK 86 import lines / 10 unique runtime paths), bun run check:runtime-cycles:strict (OK no cycles in src/main/runtime), bun test scripts/check-runtime-cycles.test.ts (3 pass), bun run scripts/check-runtime-cycles.ts --strict --root scripts/fixtures/runtime-cycles/cyclic (expected FAIL with cycle path), bun run docs:build (PASS), bun run test:fast (PASS).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Expanded maintainability guardrails by adding explicit hotspot thresholds with baseline-aware reporting and strict drift enforcement in `scripts/check-file-budgets.ts`, while preserving legacy broad-budget visibility as warnings unless explicitly requested with `--enforce-global`. Added a new runtime import-cycle guardrail (`scripts/check-runtime-cycles.ts`) with fixture-backed tests proving strict detection of an injected cycle. Wired guardrails into CI as a fail-fast step and documented contributor remediation workflows plus baseline thresholds in `docs/development.md`. Verified with guardrail commands, cycle fixture failure check, docs build, and `test:fast`.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Guardrail scripts and CI wiring merged with passing checks.
- [x] #2 Documentation updated with local commands and expected outputs.
- [x] #3 Baseline numbers recorded to justify thresholds.
<!-- DOD:END -->
