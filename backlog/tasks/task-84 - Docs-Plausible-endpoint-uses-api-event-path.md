---
id: TASK-84
title: 'Docs Plausible endpoint uses /api/event path'
status: Done
assignee: []
created_date: '2026-03-03 00:00'
updated_date: '2026-03-03 00:00'
labels: []
dependencies: []
priority: medium
ordinal: 12000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Fix VitePress docs Plausible tracker config to post to hosted worker API event endpoint instead of worker root URL.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Docs theme Plausible `endpoint` points to `https://worker.subminer.moe/api/event`.
- [x] #2 Plausible docs test asserts `/api/event` endpoint path.

<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Updated docs Plausible tracker endpoint to `https://worker.subminer.moe/api/event` and updated regression test expectation accordingly.

<!-- SECTION:FINAL_SUMMARY:END -->
