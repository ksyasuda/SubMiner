---
id: TASK-48.1
title: Add streaming CLI flag plumbing and option wiring
status: To Do
assignee: []
created_date: '2026-02-14 06:03'
updated_date: '2026-02-14 08:19'
labels:
  - stream
  - cli
dependencies: []
parent_task_id: TASK-48
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add the `-s`/`--stream` option end-to-end in SubMiner CLI and configuration handling, including defaults, help text, parsing/validation, and explicit routing so streaming mode is only enabled when requested.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Introduce `-s` short option and `--stream` long option in CLI parsing without breaking existing flags.
- [ ] #2 When set, the resulting config state reflects streaming mode enabled and is propagated to playback/session startup.
- [ ] #3 When unset, behavior remains identical to current non-streaming flows.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Superseded by TASK-51. CLI stream mode work deferred until streaming architecture is revisited.
<!-- SECTION:NOTES:END -->
