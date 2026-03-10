---
id: TASK-149
title: Cut patch release v0.5.5 for character dictionary updates and release guarding
status: Done
assignee:
  - codex
created_date: '2026-03-09 01:10'
updated_date: '2026-03-09 01:14'
labels:
  - release
  - patch
dependencies:
  - TASK-140
  - TASK-141
  - TASK-142
  - TASK-143
  - TASK-144
  - TASK-145
  - TASK-146
  - TASK-148
references:
  - package.json
  - CHANGELOG.md
  - scripts/build-changelog.ts
  - scripts/build-changelog.test.ts
  - docs/RELEASING.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Prepare and publish patch release `v0.5.5` after the failed `v0.5.4` tag by aligning package version metadata, generating committed changelog output from the pending release fragments, and hardening release validation so a future tag cannot ship with a mismatched `package.json` version.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Repository version metadata is updated to `0.5.5`.
- [x] #2 `CHANGELOG.md` contains the committed `v0.5.5` section and the consumed fragments are removed.
- [x] #3 Release validation rejects a requested release version when it differs from `package.json`.
- [x] #4 Release docs capture the required version/changelog prep before tagging.
- [x] #5 New `v0.5.5` release-prep commit and tag are pushed to `origin/main`.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Add a regression test for tagged-release/package version mismatch.
2. Update changelog validation to reject mismatched explicit release versions.
3. Bump `package.json`, generate committed `v0.5.5` changelog output, and remove consumed fragments.
4. Add a short `docs/RELEASING.md` checklist for the prep flow.
5. Run release verification, commit, tag, and push.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Added a regression test in `scripts/build-changelog.test.ts` that proves `changelog:check --version ...` rejects tag/package mismatches. Updated `scripts/build-changelog.ts` so tagged release validation now compares the explicit requested version against `package.json` before looking for pending fragments or the committed changelog section.

Bumped `package.json` from `0.5.3` to `0.5.5`, ran `bun run changelog:build --version 0.5.5 --date 2026-03-09`, and committed the generated `CHANGELOG.md` output while removing the consumed task fragments. Added `docs/RELEASING.md` with the required release-prep checklist so version bump + changelog generation happen before tagging.

Verification: `bun run changelog:lint`, `bun run changelog:check --version 0.5.5`, `bun run typecheck`, `bun run test:fast`, and `bun test scripts/build-changelog.test.ts src/release-workflow.test.ts`. `bun run format:check` still reports many unrelated pre-existing repo-wide Prettier warnings, so touched files were checked/formatted separately with `bunx prettier`.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Prepared patch release `v0.5.5` after the failed `v0.5.4` release attempt. Release metadata now matches the upcoming tag, the pending character-dictionary/overlay/plugin fragments are committed into `CHANGELOG.md`, and release validation now blocks future tag/package mismatches before publish.

Docs now include a short release checklist in `docs/RELEASING.md`. Validation passed for changelog lint/check, typecheck, targeted workflow tests, and the full fast test suite. Repo-wide Prettier remains noisy from unrelated existing files, but touched release files were formatted and verified.

<!-- SECTION:FINAL_SUMMARY:END -->
