---
id: TASK-169
title: Cut minor release v0.7.0 for stats and runtime polish
status: Done
assignee:
  - codex
created_date: '2026-03-19 17:20'
updated_date: '2026-03-19 17:31'
labels:
  - release
  - docs
  - minor
dependencies:
  - TASK-168
references:
  - package.json
  - README.md
  - docs/RELEASING.md
  - docs-site/changelog.md
  - CHANGELOG.md
  - release/release-notes.md
priority: high
ordinal: 108000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Prepare the next release cut as `v0.7.0`, keeping 0-ver semantics by rolling the accumulated stats/dashboard, launcher, overlay, and stability work into the next minor line instead of a `1.0.0` release.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Repository version metadata is updated to `0.7.0`.
- [x] #2 Root release-facing docs are refreshed for the `0.7.0` release cut.
- [x] #3 `CHANGELOG.md` and `release/release-notes.md` contain the committed `v0.7.0` section and consumed fragments are removed.
- [x] #4 Public changelog/docs surfaces reflect the new release.
- [x] #5 Release-prep verification is recorded.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Bump `package.json` to `0.7.0`.
2. Refresh release-facing docs: root `README.md`, release guide versioning note, and public docs changelog summary.
3. Run `bun run changelog:build --version 0.7.0` to commit release artifacts and consume pending fragments.
4. Run release-prep verification (`changelog`, typecheck, tests, docs build if docs-site changed).
5. Update this task with notes, verification, and final summary.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Bumped `package.json` from `0.6.5` to `0.7.0` and refreshed the root release-facing copy in `README.md` so the release prep explicitly calls out the new stats/dashboard line plus the background stats daemon commands. Updated `docs/RELEASING.md` with the repo's 0-ver versioning policy and an explicit `--date` reminder after the changelog generator initially stamped `2026-03-20` from UTC instead of the intended local release date `2026-03-19`.

Ran `bun run changelog:build --version 0.7.0`, which generated `CHANGELOG.md` and `release/release-notes.md` and removed the queued `changes/*.md` fragments for the accumulated stats, launcher, overlay, JLPT, and stability work. Added a curated `v0.7.0` summary to `docs-site/changelog.md` so the public docs changelog stays aligned with the committed root changelog while remaining user-facing.

Verification:
- `bash .agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh`
- `bun run changelog:lint`
- `bun run changelog:check --version 0.7.0`
- `bun run verify:config-example`
- `bun run typecheck`
- `bun run test:fast`
- `bun run test:env`
- `bun run build`
- `bun run docs:test`
- `bun run docs:build`
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Prepared minor release `v0.7.0` as the next 0-ver major line. Version metadata, root changelog, generated release notes, README release copy, release-guide policy, and the public docs changelog are now aligned for the release cut.

Docs update required: yes. Completed in `README.md`, `docs/RELEASING.md`, and `docs-site/changelog.md`.
Changelog fragment required: no new fragment for this task. Existing pending release fragments were consumed into the committed `v0.7.0` changelog section and `release/release-notes.md`.

Release-prep verification passed across changelog validation, config-example verification, typecheck, fast/env tests, full build, and docs-site test/build.
<!-- SECTION:FINAL_SUMMARY:END -->
