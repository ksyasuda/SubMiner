---
id: TASK-87.5
title: >-
  Dead architecture cleanup: delete unused registry and pipeline modules that
  are off the live path
status: To Do
assignee: []
created_date: '2026-03-06 03:20'
updated_date: '2026-03-06 03:21'
labels:
  - tech-debt
  - dead-code
milestone: m-0
dependencies:
  - TASK-87.1
  - TASK-87.2
references:
  - src/translators/index.ts
  - src/subsync/engines.ts
  - src/subtitle/pipeline.ts
  - src/tokenizers/index.ts
  - src/token-mergers/index.ts
  - src/core/services/subsync.ts
  - src/core/services/tokenizer.ts
documentation:
  - docs/reports/2026-02-22-task-100-dead-code-report.md
parent_task_id: TASK-87
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
The review found several modules that appear self-contained but unused from the application’s live execution paths: src/translators/index.ts, src/subsync/engines.ts, src/subtitle/pipeline.ts, src/tokenizers/index.ts, and src/token-mergers/index.ts. At the same time, the real runtime behavior is implemented elsewhere. This task should verify those modules are truly unused, remove or consolidate them, and clean up any stale exports, docs, or tests so contributors are not misled by duplicate architecture.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Each candidate module identified in the review is either removed as dead code or justified and reconnected to a real supported execution path.
- [ ] #2 Any stale exports, imports, or tests associated with the removed or consolidated modules are cleaned up so the codebase has a single obvious path for the affected behavior.
- [ ] #3 The cleanup does not regress live tokenization or subtitle sync behavior and the relevant verification commands remain green.
- [ ] #4 Contributor-facing documentation or internal notes no longer imply that removed duplicate architecture is part of the current design.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Re-verify each candidate module is off the live path by tracing imports from current runtime entrypoints before deleting anything.
2. Remove or consolidate truly dead modules and clean associated exports/imports/tests so only the supported path remains obvious.
3. Pay special attention to subtitle sync and tokenization surfaces, since duplicate architecture exists near active code.
4. Verify the relevant tokenization and subsync commands/tests still pass and update any stale docs or notes.
<!-- SECTION:PLAN:END -->
