---
id: TASK-252
title: Harden AUR publish release step against transient SSH failures
status: Done
assignee: []
created_date: '2026-03-29 23:46'
updated_date: '2026-03-31 19:37'
labels:
  - release
  - ci
  - aur
dependencies: []
priority: high
ordinal: 163500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Make tagged releases resilient when the automated AUR update hits transient SSH disconnects from GitHub-hosted runners. The GitHub Release should still complete successfully, while AUR publish should retry a few times and downgrade persistent AUR failures to warnings instead of failing the entire release workflow.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Tagged release workflow retries the AUR clone/push path with bounded backoff when AUR SSH disconnects transiently.
- [x] #2 Persistent AUR publish failure does not fail the overall tagged release workflow or block GitHub Release publication.
- [x] #3 Release documentation notes that AUR publish is best-effort and may need manual follow-up when retries are exhausted.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Updated .github/workflows/release.yml so AUR secret/configure/clone/push failures downgrade to warnings, clone/push retry three times with linear backoff, and the GitHub Release path remains green.

Documented AUR publish as best-effort in docs/RELEASING.md and added changes/253-aur-release-best-effort.md for PR changelog compliance.
<!-- SECTION:NOTES:END -->
