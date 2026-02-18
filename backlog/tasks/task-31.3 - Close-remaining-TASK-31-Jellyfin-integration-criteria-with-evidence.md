---
id: TASK-31.3
title: Close remaining TASK-31 Jellyfin integration criteria with evidence
status: To Do
assignee: []
created_date: '2026-02-18 02:51'
labels: []
dependencies:
  - TASK-31.1
  - TASK-31.2
references:
  - TASK-31
  - TASK-31.1
  - TASK-31.2
  - TASK-64
parent_task_id: TASK-31
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Drive TASK-31 to completion by collecting and documenting verification evidence for the remaining acceptance criteria (#2, #5, #6, #8), then update criterion status based on observed behavior and any explicit scope limits.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Authentication flow against a user-selected Jellyfin server is verified, including persisted/reused connection settings and token reuse behavior across restart.
- [ ] #2 Direct-play-first behavior is verified for compatible media profiles, with evidence that attempt order matches expected policy.
- [ ] #3 Transcoding fallback behavior is verified for incompatible media, including correct transcode parameter handoff to playback.
- [ ] #4 Failure-mode handling is verified for invalid credentials, token expiry, server offline, and stream/transcode error scenarios with user-visible status messaging.
- [ ] #5 TASK-31 acceptance criteria #2, #5, #6, and #8 are updated to done only when evidence is captured; otherwise each unresolved gap is explicitly documented with next action.
- [ ] #6 Project docs and/or task notes clearly summarize the final Jellyfin support boundary (working, partial, out-of-scope) for maintainers and reviewers.
<!-- AC:END -->
