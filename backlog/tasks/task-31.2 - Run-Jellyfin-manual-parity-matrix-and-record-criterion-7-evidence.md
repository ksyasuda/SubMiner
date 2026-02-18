---
id: TASK-31.2
title: Run Jellyfin manual parity matrix and record criterion-7 evidence
status: To Do
assignee: []
created_date: '2026-02-18 02:43'
updated_date: '2026-02-18 02:44'
labels: []
dependencies:
  - TASK-31.1
references:
  - TASK-31
  - TASK-31.1
  - TASK-64
documentation:
  - docs/plans/2026-02-17-jellyfin-cast-remote-playback.md
parent_task_id: TASK-31
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Validate real playback behavior against Jellyfin server media in a reproducible manual matrix, then capture evidence needed to confidently close TASK-31 acceptance criterion #7.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Manual verification covers at least one movie and one TV episode and confirms playback shows expected title/episode identity where applicable.
- [ ] #2 Manual verification confirms subtitle track selection behavior during playback, including enable/disable or track change flows where available.
- [ ] #3 Manual verification confirms audio track selection behavior during playback for media with multiple audio tracks.
- [ ] #4 Manual verification confirms resume marker behavior by stopping mid-playback and relaunching the same item.
- [ ] #5 Observed behavior, limitations, and pass/fail outcomes are documented in task notes or project docs with enough detail for reviewer validation.
- [ ] #6 TASK-31 acceptance criterion #7 is updated to done only if collected evidence satisfies all required metadata/features; otherwise remaining gaps are explicitly listed.
<!-- AC:END -->
