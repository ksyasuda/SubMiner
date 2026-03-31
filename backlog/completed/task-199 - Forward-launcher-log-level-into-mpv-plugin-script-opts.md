---
id: TASK-199
title: Forward launcher log level into mpv plugin script opts
status: Done
assignee: []
created_date: '2026-03-18 21:16'
updated_date: '2026-03-23 03:22'
labels: []
dependencies:
  - TASK-198
references:
  - /home/sudacode/projects/japanese/SubMiner/launcher/aniskip-metadata.ts
  - /home/sudacode/projects/japanese/SubMiner/launcher/mpv.ts
  - /home/sudacode/projects/japanese/SubMiner/launcher/main.test.ts
  - /home/sudacode/projects/japanese/SubMiner/launcher/aniskip-metadata.test.ts
priority: medium
ordinal: 134500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Make `subminer --log-level=debug ...` reach the mpv plugin auto-start path by forwarding the launcher log level into `--script-opts`, so plugin-started overlay and texthooker subprocesses inherit debug logging.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launcher mpv playback includes `subminer-log_level=<level>` in `--script-opts` when a non-info CLI log level is used.
- [x] #2 Detached idle mpv launch uses the same script-opt forwarding.
- [x] #3 Regression tests cover launcher script-opt forwarding.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a failing launcher regression test that captures mpv argv and expects `subminer-log_level=debug` inside `--script-opts`.
2. Extend the shared script-opt builder to accept launcher log level and emit `subminer-log_level` for non-info runs.
3. Reuse that builder in both normal mpv playback and detached idle mpv launch.
4. Run focused launcher tests and launcher-plugin verification.
<!-- SECTION:PLAN:END -->

## Outcome

<!-- SECTION:OUTCOME:BEGIN -->
Forwarded launcher log level into mpv plugin script opts via the shared builder and reused that builder for idle mpv launch. `subminer --log-level=debug ...` now gives the plugin `opts.log_level=debug`, so auto-started overlay and texthooker subprocesses include `--log-level debug` and the tokenizer timing logs can actually appear in the app log.
<!-- SECTION:OUTCOME:END -->
