---
id: TASK-259
title: Fix integrated --start --texthooker startup skipping texthooker server
status: Done
assignee: []
created_date: '2026-03-30 06:48'
updated_date: '2026-03-30 06:56'
labels:
  - bug
  - texthooker
  - startup
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Integrated overlay startup with `--start --texthooker` currently takes the minimal-startup path because startup mode flags treat any `args.texthooker` as texthooker-only. That skips app-ready texthooker service startup, so no server binds on port 5174 during normal SubMiner playback launches.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `--start --texthooker` uses full app-ready startup instead of minimal texthooker-only startup
- [x] #2 Integrated playback launch starts the texthooker server on the configured/default port
- [x] #3 Regression tests cover the startup-mode classification and integrated startup behavior
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Narrowed texthooker-only startup classification so integrated `--start --texthooker` no longer takes the minimal-startup path. Added CLI arg regression coverage, rebuilt the AppImage, installed it to `~/.local/bin/SubMiner.AppImage` with a timestamped backup, restarted against `/tmp/subminer-socket`, and verified listeners on 5174/6677/6678 plus browser connection state `Connected with ws://127.0.0.1:6678`.
<!-- SECTION:FINAL_SUMMARY:END -->
