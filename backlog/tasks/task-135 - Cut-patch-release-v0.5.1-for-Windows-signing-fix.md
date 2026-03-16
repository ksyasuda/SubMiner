---
id: TASK-135
title: Cut patch release v0.5.1 for Windows signing fix
status: Done
assignee:
  - codex
created_date: '2026-03-08 20:24'
updated_date: '2026-03-16 05:13'
labels:
  - release
  - patch
dependencies:
  - TASK-134
references:
  - package.json
  - CHANGELOG.md
  - release/release-notes.md
priority: high
ordinal: 50500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Publish a patch release from the workflow-signing fix on `main` by bumping the app version, generating the committed changelog artifacts for the new version, and pushing a new `v0.5.1` tag instead of rewriting the failed `v0.5.0` tag.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Repository version metadata is updated to `0.5.1`.
- [ ] #2 `CHANGELOG.md` and `release/release-notes.md` contain the committed `v0.5.1` section and released fragments are removed.
- [ ] #3 New `v0.5.1` commit and tag are pushed to `origin`.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Bump the package version to `0.5.1`.
2. Run the changelog builder so `CHANGELOG.md`/`release-notes.md` match the release workflow contract.
3. Run the relevant verification commands.
4. Commit the release-prep changes, create `v0.5.1`, and push both commit and tag.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Bumped `package.json` from `0.5.0` to `0.5.1`, then ran `bun run changelog:build` so the committed release artifacts match the release workflow contract. That prepended the `v0.5.1` section to `CHANGELOG.md`, regenerated `release/release-notes.md`, and removed the consumed changelog fragments from `changes/`.

Verification before tagging: `bun run changelog:lint`, `bun run changelog:check --version 0.5.1`, `bun run typecheck`, and `bun run test:fast`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Prepared patch release `v0.5.1` from the signing-workflow fix on `main` instead of rewriting the failed `v0.5.0` tag. Repository version metadata, changelog, and committed release notes are all aligned with the new release tag, and the consumed changelog fragments were removed.

Validation: `bun run changelog:lint`, `bun run changelog:check --version 0.5.1`, `bun run typecheck`, `bun run test:fast`.
<!-- SECTION:FINAL_SUMMARY:END -->
