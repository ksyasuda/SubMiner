---
id: TASK-118
title: Fix GitHub release workflow publish step failure
status: Done
assignee:
  - Codex
created_date: '2026-03-08 03:34'
updated_date: '2026-03-08 03:38'
labels:
  - ci
  - release
  - github-actions
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/.github/workflows/release.yml
  - 'https://github.com/ksyasuda/SubMiner/actions/runs/22812335927'
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

The GitHub Actions Release workflow fails during the Publish Release step for tag releases because the gh CLI invocation passes invalid arguments when creating or editing the GitHub release. Restore successful release publication for tagged builds without changing unrelated release packaging behavior.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Tagged Release workflow completes the Publish Release step without gh CLI argument errors.
- [x] #2 Release workflow still creates or updates the GitHub release as a non-prerelease for normal version tags.
- [x] #3 A regression check covers the publish command shape or workflow behavior that caused this failure.
- [x] #4 Any release workflow behavior change is documented in repository docs or workflow comments if needed.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Add a targeted regression test for .github/workflows/release.yml that fails if the publish step passes an argument to the gh --prerelease boolean flag or otherwise omits explicit non-prerelease behavior.
2. Run the targeted test to confirm the current workflow fails for the expected reason.
3. Patch the Publish Release step in .github/workflows/release.yml to remove the invalid gh CLI usage while preserving non-prerelease release creation/update behavior.
4. Re-run the targeted regression test and any relevant lightweight verification, then record results in task notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Identified root cause from GitHub Actions run 22812335927: Publish Release failed with `accepts 1 arg(s), received 2` because the workflow passed a value to gh's boolean prerelease flag.

Added a workflow comment clarifying that omitting the prerelease flag keeps normal releases as non-prerelease releases.

Added src/release-workflow.test.ts and wired it into `bun run test:fast` so CI catches the invalid workflow shape before the next tag.

Verification: `bun test src/release-workflow.test.ts`, `bun run typecheck`, and `bun run test:fast` all passed locally.

Code-review pass found no issues; remaining caveat is that prerelease tag semantics are still not modeled for tags like `v1.0.0-beta.1`, which is outside this fix scope.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Fixed the GitHub Actions release publish step so tagged releases no longer fail on invalid gh CLI usage. The workflow now omits the prerelease flag when creating or editing normal releases, which preserves existing non-prerelease behavior and avoids the `accepts 1 arg(s), received 2` failure seen in run 22812335927.

Added a small regression test that reads `.github/workflows/release.yml` and asserts the publish step does not set the prerelease flag, then included that test in `bun run test:fast` so the main verification lane catches this class of workflow regression before the next release.

Validation run locally: `bun test src/release-workflow.test.ts`, `bun run typecheck`, and `bun run test:fast`. Residual risk: prerelease-tag semantics remain unchanged for tags such as `v1.0.0-beta.1`; this fix is intentionally scoped to restoring normal tagged release publication.

<!-- SECTION:FINAL_SUMMARY:END -->
