---
id: TASK-87
title: >-
  Codebase health: harden verification and retire dead architecture identified
  in the March 2026 review
status: To Do
assignee: []
created_date: '2026-03-06 03:19'
updated_date: '2026-03-06 03:20'
labels:
  - tech-debt
  - tests
  - maintainability
milestone: m-0
dependencies: []
references:
  - package.json
  - README.md
  - src/main.ts
  - src/anki-integration.ts
  - src/core/services/immersion-tracker-service.test.ts
  - src/translators/index.ts
  - src/subsync/engines.ts
  - src/subtitle/pipeline.ts
documentation:
  - docs/reports/2026-02-22-task-100-dead-code-report.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Track the remediation work from the March 6, 2026 code review. The review found that the default test gate only exercises 53 of 241 test files, the dedicated subtitle test lane is a no-op, SQLite-backed immersion tracking tests are conditionally skipped in the standard Bun run, src/main.ts still contains a large dead-symbol backlog, several registry/pipeline modules appear unreferenced from live execution paths, and src/anki-integration.ts remains an oversized orchestration file. This parent task should coordinate a safe sequence: improve verification first, then remove dead code and continue decomposition with good test coverage in place.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Child tasks are created for each remediation workstream with explicit dependencies and enough context for an isolated agent to execute them.
- [ ] #2 The parent task records the recommended sequencing and parallelization strategy so replacement agents can resume without conversation history.
- [ ] #3 Completion of the parent task leaves the repository with a materially more trustworthy test gate, less dead architecture, and clearer ownership boundaries for the main runtime and Anki integration surfaces.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Recommended sequencing:
1. Run TASK-87.1, TASK-87.2, TASK-87.3, and TASK-87.7 first. These are the safety-net and tooling tasks and can largely proceed in parallel.
2. Start TASK-87.4 once TASK-87.1 lands so src/main.ts cleanup happens under a more trustworthy test matrix.
3. Start TASK-87.5 after TASK-87.1 and TASK-87.2 so dead subsync/pipeline cleanup happens with stronger subtitle and runtime verification.
4. Start TASK-87.6 after TASK-87.1 so Anki refactors happen with broader default coverage in place.
5. Keep PRs focused: do not combine verification work with architectural cleanup unless a narrow dependency requires it.

Parallelization guidance:
- Wave 1 parallel: TASK-87.1, TASK-87.2, TASK-87.3, TASK-87.7
- Wave 2 parallel: TASK-87.4, TASK-87.5, TASK-87.6

Shared review context to restate in child tasks:
- Standard test scripts currently reference only 53 unique test files out of 241 discovered test and type-test files under src/ and launcher/.
- test:subtitle is currently a placeholder echo even though subtitle sync is a user-facing feature.
- SQLite-backed immersion tracker tests are conditionally skipped in the standard Bun run.
- src/main.ts trips many noUnusedLocals/noUnusedParameters diagnostics.
- src/translators/index.ts, src/subsync/engines.ts, src/subtitle/pipeline.ts, src/tokenizers/index.ts, and src/token-mergers/index.ts appeared unreferenced during review and must be re-verified before deletion.
<!-- SECTION:PLAN:END -->
