---
id: TASK-137
title: Cut patch release v0.5.2 for SignPath artifact config pinning
status: Done
assignee:
  - codex
created_date: '2026-03-08 20:44'
updated_date: '2026-03-18 05:28'
labels:
  - release
  - patch
dependencies:
  - TASK-136
references:
  - package.json
  - CHANGELOG.md
  - release/release-notes.md
priority: high
ordinal: 50500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Publish a patch release from the SignPath artifact-configuration pinning change by bumping the app version, generating the committed changelog artifacts for the new version, and pushing a new `v0.5.2` tag.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Repository version metadata is updated to `0.5.2`.
- [ ] #2 `CHANGELOG.md` and `release/release-notes.md` contain the committed `v0.5.2` section and consumed fragments are removed.
- [ ] #3 New `v0.5.2` commit and tag are pushed to `origin`.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add the release fragment for the SignPath configuration pinning change.
2. Bump `package.json` to `0.5.2` and run the changelog builder.
3. Run changelog/typecheck/test verification.
4. Commit the release-prep change set, create `v0.5.2`, and push commit plus tag.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Bumped `package.json` from `0.5.1` to `0.5.2`, ran `bun run changelog:build`, and committed the generated release artifacts. That prepended the `v0.5.2` section to `CHANGELOG.md`, regenerated `release/release-notes.md`, and removed the consumed `changes/signpath-artifact-config-pin.md` fragment.

Verification before tagging: `bun run changelog:lint`, `bun run changelog:check --version 0.5.2`, `bun run typecheck`, and `bun run test:fast`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Prepared patch release `v0.5.2` so the explicit SignPath artifact-configuration pin ships on a fresh release tag. Version metadata, committed changelog artifacts, and release notes are aligned with the new patch version.

Validation: `bun run changelog:lint`, `bun run changelog:check --version 0.5.2`, `bun run typecheck`, `bun run test:fast`.
<!-- SECTION:FINAL_SUMMARY:END -->
