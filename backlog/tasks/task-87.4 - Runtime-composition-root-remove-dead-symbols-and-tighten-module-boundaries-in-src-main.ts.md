---
id: TASK-87.4
title: >-
  Runtime composition root: remove dead symbols and tighten module boundaries in
  src/main.ts
status: Done
assignee: []
created_date: '2026-03-06 03:19'
updated_date: '2026-03-06 18:10'
labels:
  - tech-debt
  - runtime
  - maintainability
milestone: m-0
dependencies:
  - TASK-87.1
references:
  - src/main.ts
  - src/main/runtime
  - package.json
documentation:
  - docs/reports/2026-02-22-task-100-dead-code-report.md
parent_task_id: TASK-87
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

A noUnusedLocals/noUnusedParameters compile pass reports a large concentration of dead imports and dead locals in src/main.ts. The file is also far beyond the repo’s preferred size guideline, which makes the runtime composition root difficult to review and easy to break. This task should remove confirmed dead symbols, continue extracting coherent slices where that improves readability, and leave the entrypoint materially easier to understand without changing behavior.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 src/main.ts no longer emits dead-symbol diagnostics under a noUnusedLocals/noUnusedParameters compile pass for the areas touched by this cleanup.
- [x] #2 Unused imports, destructured values, and stale locals identified in the current composition root are removed or relocated without behavior changes.
- [x] #3 The resulting composition root has clearer ownership boundaries for at least one runtime slice that is currently buried in the monolith.
- [x] #4 Relevant runtime and startup verification commands pass after the cleanup, and any command changes are documented if needed.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Re-run the noUnusedLocals/noUnusedParameters compile pass and capture the src/main.ts diagnostics cluster before editing.
2. Remove dead imports, destructured values, and stale locals in small reviewable slices; extract a coherent helper/module only where that materially reduces coupling.
3. Keep changes behavior-preserving and avoid mixing unrelated cleanup outside src/main.ts unless required to compile.
4. Verify with the updated runtime/startup test commands from TASK-87.1 plus a noUnused compile pass.
<!-- SECTION:PLAN:END -->

## Completion Notes

- Removed the dead import/destructure backlog from `src/main.ts` and deleted stale wrapper seams that no longer owned runtime behavior after the composer/runtime extractions.
- Tightened module boundaries so the composition root depends on the composed/public runtime surfaces it actually uses instead of retaining unused lower-level domain factory symbols.
- Cleared the remaining strict `noUnusedLocals`/`noUnusedParameters` failures in nearby touched files required for a clean repo-wide pass: `launcher/commands/playback-command.ts`, `src/anki-integration.ts`, `src/anki-integration/field-grouping-workflow.ts`, `src/core/services/tokenizer/yomitan-parser-runtime.test.ts`, and `src/main/runtime/composers/composer-contracts.type-test.ts`.
- Verification:
  - `bunx tsc --noEmit -p tsconfig.typecheck.json --noUnusedLocals --noUnusedParameters --pretty false`
  - `bun run test:fast`
- Commit: `e659b5d` (`refactor(runtime): remove dead symbols from composition roots`)
