---
id: TASK-31.1
title: Verify Jellyfin playback metadata parity with automated coverage
status: To Do
assignee: []
created_date: '2026-02-18 02:43'
labels: []
dependencies: []
references:
  - TASK-31
  - TASK-64
parent_task_id: TASK-31
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Establish objective pass/fail evidence that Jellyfin playback preserves metadata and media-feature parity needed for TASK-31 acceptance criterion #7, so completion is based on repeatable test coverage rather than ad-hoc checks.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Automated test coverage verifies Jellyfin playback launch preserves title and episodic identity metadata when provided by server data.
- [ ] #2 Automated test coverage verifies subtitle and audio track selection behavior is preserved through playback launch and control paths.
- [ ] #3 Automated test coverage verifies resume position/marker behavior is preserved for partially watched items.
- [ ] #4 Tests include at least one edge scenario with incomplete metadata or missing track info and assert graceful behavior.
- [ ] #5 Project test suite passes with the new/updated Jellyfin parity tests included.
- [ ] #6 Test expectations and scope are documented in repository docs or task notes so future contributors can reproduce verification intent.
<!-- AC:END -->
