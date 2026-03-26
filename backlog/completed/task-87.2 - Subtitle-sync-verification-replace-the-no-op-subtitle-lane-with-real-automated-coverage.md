---
id: TASK-87.2
title: >-
  Subtitle sync verification: replace the no-op subtitle lane with real
  automated coverage
status: Done
assignee:
  - Kyle Yasuda
created_date: '2026-03-06 03:19'
updated_date: '2026-03-16 05:13'
labels:
  - tests
  - subsync
milestone: m-0
dependencies: []
references:
  - package.json
  - README.md
  - src/core/services/subsync.ts
  - src/core/services/subsync.test.ts
  - src/subsync/utils.ts
documentation:
  - docs/reports/2026-02-22-task-100-dead-code-report.md
parent_task_id: TASK-87
priority: high
ordinal: 92500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
SubMiner advertises subtitle syncing with alass and ffsubsync, but the dedicated test:subtitle command currently does not run any tests. There is already lower-level coverage in src/core/services/subsync.test.ts, but the test matrix and contributor-facing commands do not reflect that reality. This task should replace the no-op lane with real verification, align scripts with the existing subsync test surface, and make the user-facing docs honest about how subtitle sync is verified.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 The test:subtitle entrypoint runs real automated verification instead of echoing a placeholder message.
- [x] #2 The subtitle verification lane covers both alass and ffsubsync behavior, including at least one non-happy-path scenario relevant to current functionality.
- [x] #3 Contributor-facing documentation points to the real subtitle verification command and no longer implies a dedicated test lane exists when it does not.
- [x] #4 The resulting verification strategy integrates cleanly with the repository-wide test matrix without duplicating or hiding existing subsync coverage.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Plan of record:

1. Replace the placeholder package-script lane with a real `test:subtitle:src` command that runs the maintained subtitle-sync tests directly (`src/core/services/subsync.test.ts` and `src/subsync/utils.test.ts`), and point `test:subtitle` at that lane instead of build+echo behavior.
2. Add one focused ffsubsync non-happy-path test in `src/core/services/subsync.test.ts` so the dedicated lane explicitly covers both engines plus failure propagation relevant to current functionality.
3. Update `README.md` contributor guidance to name `bun run test:subtitle` as the subtitle verification command and explain that it reuses the maintained subsync tests already included in broader core coverage.
4. Verify the final strategy by running `bun run test:subtitle` and `bun run test:core:src` so the dedicated lane stays aligned with the repository-wide matrix instead of creating a divergent hidden suite.

Detailed execution plan saved at `docs/plans/2026-03-06-subtitle-sync-verification.md`.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Reviewed task references and current subtitle verification surface. Existing coverage already lives in `src/core/services/subsync.test.ts` and `src/subsync/utils.test.ts`; `test:subtitle` is still a placeholder build+echo wrapper. The referenced report `docs/reports/2026-02-22-task-100-dead-code-report.md` is not present in the workspace, so planning used the task body plus repository state instead.

Implementation plan written and saved to `docs/plans/2026-03-06-subtitle-sync-verification.md`. Proceeding with execution per the task request.

Replaced the placeholder subtitle lane with `test:subtitle:src` in `package.json`, pointing `test:subtitle` directly at `src/core/services/subsync.test.ts` and `src/subsync/utils.test.ts` instead of build+echo behavior.

Added explicit ffsubsync failure-path coverage in `src/core/services/subsync.test.ts`, asserting non-zero command failures surface detailed `ffsubsync synchronization failed` messaging alongside existing alass coverage.

Updated `README.md` verification guidance to point contributors at `bun run test:subtitle` and explain that the lane reuses the maintained subsync tests already included in `bun run test:core`.

Verification: `bun run test:subtitle` passed (15 tests across 2 files). `bun run test:core:src` also passed (373 pass, 6 skip, 0 fail), confirming the dedicated subtitle lane stays aligned with the broader matrix.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented a real subtitle verification lane by replacing the placeholder `test:subtitle` build+echo flow with a source-level `test:subtitle:src` command that runs the maintained subtitle-sync tests directly from `src/core/services/subsync.test.ts` and `src/subsync/utils.test.ts`. This keeps subtitle verification explicit for contributors while still reusing the same maintained test surface already covered by `test:core`.

Expanded subtitle-sync coverage with an explicit ffsubsync failure-path test so the dedicated lane now exercises both engines plus a user-visible non-happy path. Updated `README.md` to document `bun run test:subtitle` as the contributor-facing subtitle verification command and to explain its relationship to the broader core suite.

Verification run:

- `bun run test:subtitle`
- `bun run test:core:src`

Notes:

- The task reference `docs/reports/2026-02-22-task-100-dead-code-report.md` was not present in the workspace during execution, so implementation used the task body and live repository state as the source of truth.
<!-- SECTION:FINAL_SUMMARY:END -->
