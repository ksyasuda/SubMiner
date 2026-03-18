---
id: TASK-136
title: Pin SignPath artifact configuration in release workflow
status: Done
assignee:
  - codex
created_date: '2026-03-08 20:41'
updated_date: '2026-03-18 05:28'
labels:
  - ci
  - release
  - windows
  - signing
dependencies:
  - TASK-134
references:
  - .github/workflows/release.yml
  - build/signpath-windows-artifact-config.xml
  - src/release-workflow.test.ts
priority: high
ordinal: 49500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
The Windows release workflow currently relies on the default SignPath artifact configuration configured in the SignPath UI. Pin the workflow to an explicit artifact-configuration slug so the checked-in signing configuration and CI behavior stay deterministic across future SignPath project changes.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 The Windows release workflow validates a dedicated SignPath artifact-configuration secret/input.
- [ ] #2 Every SignPath submission attempt passes `artifact-configuration-slug`.
- [ ] #3 Regression coverage fails if the explicit SignPath artifact-configuration binding is removed.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a failing workflow regression test for the explicit SignPath artifact-configuration slug.
2. Patch the Windows signing secret validation and SignPath action inputs to require the slug.
3. Run targeted release-workflow verification plus the standard fast lane.
4. Cut a new patch release so the tag-triggered release workflow runs with the pinned SignPath configuration.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added regression coverage in `src/release-workflow.test.ts` for an explicit SignPath artifact-configuration slug so the release workflow test now fails if the slug validation or action input is removed.

Patched `.github/workflows/release.yml` so Windows signing now requires `SIGNPATH_ARTIFACT_CONFIGURATION_SLUG` during secret validation and passes `artifact-configuration-slug: ${{ secrets.SIGNPATH_ARTIFACT_CONFIGURATION_SLUG }}` on every SignPath submission attempt.

Verification: `bun test src/release-workflow.test.ts`, `bun run typecheck`, `bun run test:fast`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
The release workflow is now pinned to an explicit SignPath artifact configuration instead of relying on whichever SignPath artifact config is marked default in the UI. Windows signing secret validation fails fast if `SIGNPATH_ARTIFACT_CONFIGURATION_SLUG` is missing, and every SignPath submission attempt now includes the pinned slug.

Validation: `bun test src/release-workflow.test.ts`, `bun run typecheck`, `bun run test:fast`.
<!-- SECTION:FINAL_SUMMARY:END -->
