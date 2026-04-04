---
id: TASK-278
title: Prepare patch release 0.11.1
status: Done
assignee:
  - '@codex'
created_date: '2026-04-04 07:16'
updated_date: '2026-04-04 07:41'
labels:
  - release
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Bump SubMiner from 0.11.0 to 0.11.1, run the local release checklist, and confirm whether the branch is ready for tagging/publishing.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 package.json is bumped to 0.11.1 without overwriting unrelated local metadata edits.
- [x] #2 Release prep checks are run and summarized, including changelog/build verification and local build/test gates.
- [x] #3 Any remaining release blockers are called out explicitly.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Generate release metadata with `bun run changelog:build --version 0.11.1 --date 2026-04-04`.
2. Review the resulting `CHANGELOG.md` and `release/release-notes.md` content for the 0.11.1 section.
3. Rerun the documented local release gate: `bun run changelog:check --version 0.11.1`, `bun run verify:config-example`, `bun run typecheck`, `bun run test:fast`, `bun run test:env`, and `bun run build`.
4. Record results, remaining blockers, and ready-for-tagging status in the task notes/final summary.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Bumped package.json from 0.11.0 to 0.11.1 while preserving the existing local productName/desktopName edits.

Fixed the Linux shortcut changelog fragment metadata so changelog lint now passes and the previously failing CI cause is addressed locally.

Release checks run: bun run changelog:lint, bun run changelog:check --version 0.11.1, bun run verify:config-example, bun run typecheck, bun run test:fast, bun run test:env, bun run build. All passed except changelog:check, which is expected until pending fragments are built into CHANGELOG.md/release/release-notes.md.

Current release blocker: pending changelog fragments /changes/2026-04-04-linux-app-id-metadata.md and /changes/2026-04-04-linux-shortcut-portal-regression.md still need bun run changelog:build --version 0.11.1 --date 2026-04-04 before tagging.

Reopened to complete the remaining release-prep work: generate changelog/release notes artifacts, rerun the documented release gate, and confirm ready-for-tagging status.

Completed the remaining release-prep step with `bun run changelog:build --version 0.11.1 --date 2026-04-04`, which prepended `CHANGELOG.md`, generated `release/release-notes.md`, and removed the two pending `changes/*.md` fragments.

Reran the release gate after changelog generation: `bun run changelog:check --version 0.11.1`, `bun run verify:config-example`, `bun run typecheck`, `bun run test:fast`, `bun run test:env`, and `bun run build`; all passed.

Extra confidence checks also passed: `bun run changelog:lint`, `bun run test:smoke:dist`, and `gh run list --workflow CI --limit 5` shows the two latest `main` CI runs succeeded on 2026-04-04 after the earlier pre-fix failure.

`release/release-notes.md` is intentionally generated under the gitignored `release/` directory, so readiness evidence in git status is `CHANGELOG.md` plus deletion of the released change fragments.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Finished the 0.11.1 release prep by generating the changelog artifacts from the pending Linux fix fragments and re-running the full documented local release gate. `CHANGELOG.md` now contains the `v0.11.1 (2026-04-04)` section, `release/release-notes.md` was regenerated in the ignored `release/` directory, and the released `changes/2026-04-04-linux-app-id-metadata.md` plus `changes/2026-04-04-linux-shortcut-portal-regression.md` fragments were removed.

Verification results: `bun run changelog:check --version 0.11.1`, `bun run verify:config-example`, `bun run typecheck`, `bun run test:fast`, `bun run test:env`, `bun run build`, `bun run changelog:lint`, and `bun run test:smoke:dist` all passed locally. Remote CI also looks green for the latest release-prep head: `gh run list --workflow CI --limit 5` showed successful `main` runs for `chore: prep 0.11.1 release` and `Change demo image link to GitHub asset` on 2026-04-04, with the earlier `fix: restore linux modal shortcuts` failure already superseded by later green runs.

Remaining manual release step: commit the generated release-prep changes if desired, tag `v0.11.1`, and push the commit plus tag when ready.
<!-- SECTION:FINAL_SUMMARY:END -->
