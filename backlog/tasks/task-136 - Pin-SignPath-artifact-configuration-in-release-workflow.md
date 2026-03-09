---
id: TASK-136
title: Pin SignPath artifact configuration in release workflow
status: In Progress
assignee:
  - codex
created_date: '2026-03-08 20:41'
updated_date: '2026-03-08 20:41'
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
