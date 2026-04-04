---
id: TASK-278
title: Prepare patch release 0.11.1
status: Done
assignee:
  - '@codex'
created_date: '2026-04-04 07:16'
updated_date: '2026-04-04 07:19'
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

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Bumped package.json from 0.11.0 to 0.11.1 while preserving the existing local productName/desktopName edits.

Fixed the Linux shortcut changelog fragment metadata so changelog lint now passes and the previously failing CI cause is addressed locally.

Release checks run: bun run changelog:lint, bun run changelog:check --version 0.11.1, bun run verify:config-example, bun run typecheck, bun run test:fast, bun run test:env, bun run build. All passed except changelog:check, which is expected until pending fragments are built into CHANGELOG.md/release/release-notes.md.

Current release blocker: pending changelog fragments /changes/2026-04-04-linux-app-id-metadata.md and /changes/2026-04-04-linux-shortcut-portal-regression.md still need bun run changelog:build --version 0.11.1 --date 2026-04-04 before tagging.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Prepared local patch release version 0.11.1 and ran the release gate. changelog:lint, verify:config-example, typecheck, test:fast, test:env, and build all passed. changelog:check still fails because the pending changelog fragments have not been built into CHANGELOG.md/release notes yet; run bun run changelog:build --version 0.11.1 --date 2026-04-04 as the remaining pre-tag step.
<!-- SECTION:FINAL_SUMMARY:END -->
