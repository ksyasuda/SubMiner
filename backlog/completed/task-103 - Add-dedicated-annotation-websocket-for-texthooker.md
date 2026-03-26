---
id: TASK-103
title: Add dedicated annotation websocket for texthooker
status: Done
assignee:
  - codex
created_date: '2026-03-07 02:20'
updated_date: '2026-03-16 05:13'
labels:
  - texthooker
  - websocket
  - subtitle
dependencies: []
priority: medium
ordinal: 73500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add a separate annotated subtitle websocket for bundled texthooker so token/JLPT/frequency markup is available on a stable dedicated port even when the regular websocket is in `auto` mode and skipped because `mpv_websocket` is installed.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Regular `websocket.enabled: "auto"` behavior remains unchanged and still skips the regular websocket when `mpv_websocket` is installed.
- [x] #2 A separate `annotationWebsocket` config controls an independent annotated websocket with default port `6678`.
- [x] #3 Bundled texthooker is pointed at the annotation websocket when it is enabled.
- [x] #4 Focused regression tests cover config parsing, startup wiring, and texthooker bootstrap injection.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added `annotationWebsocket.enabled`/`annotationWebsocket.port` with defaults of `true`/`6678`, started that websocket independently from the regular auto-managed websocket, and injected the bundled texthooker websocket URL so it connects to the annotation feed by default.

Also added focused regression coverage and regenerated the checked-in config examples.
<!-- SECTION:FINAL_SUMMARY:END -->
