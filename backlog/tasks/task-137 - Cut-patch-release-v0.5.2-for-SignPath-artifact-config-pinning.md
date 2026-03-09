---
id: TASK-137
title: Cut patch release v0.5.2 for SignPath artifact config pinning
status: In Progress
assignee:
  - codex
created_date: '2026-03-08 20:44'
updated_date: '2026-03-08 20:44'
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
