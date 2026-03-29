---
id: TASK-245
title: Cut minor release v0.10.0 for docs and release prep
status: Done
assignee:
  - '@codex'
created_date: '2026-03-29 08:10'
updated_date: '2026-03-29 08:13'
labels:
  - release
  - docs
  - minor
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/package.json
  - /home/sudacode/projects/japanese/SubMiner/README.md
  - /home/sudacode/projects/japanese/SubMiner/docs/RELEASING.md
  - /home/sudacode/projects/japanese/SubMiner/docs/README.md
  - /home/sudacode/projects/japanese/SubMiner/docs-site/changelog.md
  - /home/sudacode/projects/japanese/SubMiner/CHANGELOG.md
  - /home/sudacode/projects/japanese/SubMiner/release/release-notes.md
priority: high
ordinal: 54850
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Prepare the next 0-ver minor release cut as `v0.10.0`, keeping release-facing docs, backlog, and changelog artifacts aligned, then run the release-prep verification gate.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Repository version metadata is updated to `0.10.0`.
- [x] #2 Release-facing docs and public changelog surfaces are aligned for the `v0.10.0` cut.
- [x] #3 `CHANGELOG.md` and `release/release-notes.md` contain the committed `v0.10.0` section and any consumed fragments are removed.
- [x] #4 Release-prep verification passes for changelog, config example, typecheck, tests, and build.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Completed:
- Bumped `package.json` from `0.9.3` to `0.10.0`.
- Ran `bun run changelog:build --version 0.10.0 --date 2026-03-29`, which generated `CHANGELOG.md` and `release/release-notes.md` and removed the queued `changes/*.md` fragments.
- Updated `docs-site/changelog.md` with the public-facing `v0.10.0` summary.

Verification:
- `bun run changelog:lint`
- `bun run changelog:check --version 0.10.0`
- `bun run verify:config-example`
- `bun run typecheck`
- `bunx bun@1.3.5 run test:fast`
- `bunx bun@1.3.5 run test:env`
- `bunx bun@1.3.5 run build`
- `bunx bun@1.3.5 run docs:test`
- `bunx bun@1.3.5 run docs:build`

Notes:
- The local `bun` binary is `1.3.11`, which tripped Bun's nested `node:test` handling in `test:fast`; rerunning with the repo-pinned `bun@1.3.5` cleared the issue.
- No README content change was necessary for this cut.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Prepared the `v0.10.0` release cut locally. Bumped `package.json`, generated committed root changelog and release notes, updated the public docs changelog summary, and verified the release gate with the repo-pinned Bun `1.3.5` runtime. The release prep is green and ready for tagging/publishing when desired.
<!-- SECTION:FINAL_SUMMARY:END -->
