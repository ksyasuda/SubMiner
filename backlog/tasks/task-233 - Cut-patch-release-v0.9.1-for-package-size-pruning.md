---
id: TASK-233
title: Cut patch release v0.9.1 for package size pruning
status: Done
assignee:
  - '@codex'
created_date: '2026-03-24 12:40'
updated_date: '2026-03-24 12:55'
labels:
  - release
  - patch
dependencies:
  - TASK-232
references:
  - /home/sudacode/projects/japanese/SubMiner/package.json
  - /home/sudacode/projects/japanese/SubMiner/CHANGELOG.md
  - /home/sudacode/projects/japanese/SubMiner/release/release-notes.md
  - /home/sudacode/projects/japanese/SubMiner/docs-site/changelog.md
priority: high
ordinal: 54800
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Publish a patch release for the packaging-size cleanup by bumping the app version to `0.9.1`, generating committed release metadata, and keeping release-facing docs/changelog surfaces aligned.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Repository version metadata is updated to `0.9.1`.
- [x] #2 `CHANGELOG.md`, `release/release-notes.md`, and `docs-site/changelog.md` contain the committed `v0.9.1` release line and the consumed fragment is removed.
- [x] #3 Release-readiness verification passes for changelog, docs, tests, and build lanes.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Completed:
- Bumped `package.json` to `0.9.1`.
- Ran `bun run changelog:build --version 0.9.1 --date 2026-03-24`, which generated `CHANGELOG.md` + `release/release-notes.md` and consumed both pending release fragments.
- Synced `docs-site/changelog.md` with the generated `v0.9.1` release line.
- Confirmed no additional README/docs wording changes were needed beyond changelog surfaces.

Verification:
- `bun run changelog:lint`
- `bun run changelog:check --version 0.9.1`
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

Prepared patch release `v0.9.1` locally. Version metadata, committed changelog artifacts, release notes, and docs-site changelog are aligned, and the release gate is green. Pending manual release actions are the release-prep commit, `git tag v0.9.1`, and push/tag publication.

<!-- SECTION:FINAL_SUMMARY:END -->
