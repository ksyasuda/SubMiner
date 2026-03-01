---
id: TASK-73
title: 'MPV plugin: split into modules and optimize startup/command runtime'
status: Done
assignee: []
created_date: '2026-02-28 20:50'
updated_date: '2026-02-28 22:36'
labels: []
dependencies: []
references:
  - plugin/subminer/main.lua
  - plugin/subminer/bootstrap.lua
  - plugin/subminer/process.lua
  - plugin/subminer/aniskip.lua
  - plugin/subminer/environment.lua
  - plugin/subminer/lifecycle.lua
  - plugin/subminer/messages.lua
  - plugin/subminer/ui.lua
  - plugin/subminer/hover.lua
  - plugin/subminer/options.lua
  - plugin/subminer/state.lua
  - plugin/subminer.conf
  - scripts/test-plugin-start-gate.lua
  - scripts/test-plugin-process-start-retries.lua
  - launcher/commands/playback-command.ts
  - launcher/mpv.ts
  - launcher/mpv.test.ts
  - launcher/smoke.e2e.test.ts
  - Makefile
  - package.json
  - docs/mpv-plugin.md
  - docs/installation.md
  - docs/architecture.md
  - README.md
priority: medium
ordinal: 4000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Scope: Replace monolithic `plugin/subminer.lua` with modular plugin runtime; optimize command execution paths; align install/docs/tests; fix launcher smoke instability.

Delivered behavior:

- Full plugin cutover to `plugin/subminer/main.lua` + module directory (no runtime compatibility shim with old monolith file).
- Process/control command path moved toward async subprocess usage for non-start actions (`stop`, `toggle`, `settings`, restart stop leg), reducing synchronous blocking in mpv script runtime.
- AniSkip path guarded: lookup runs only in SubMiner context (launcher metadata, explicit script-message refresh, or detected running app), instead of every opened file.
- AniSkip lookup pipeline moved to async subprocess calls (no sync `ps`/`curl` on `file-loaded`) with deferred fetch after auto-start and session-level MAL/title/payload caching.
- Startup/runtime loading updated with lazy module initialization via bootstrap proxies.
- Plugin install flow updated to copy `plugin/subminer/` directory and remove legacy `~/.config/mpv/scripts/subminer.lua` file.
- Added plugin gate script wiring to package scripts (`test:plugin:src`) and launcher test flow.
- Smoke tests stabilized across sandbox environments where UNIX socket bind can return `EPERM` while preserving normal-path assertions.
- Playback command cleanup race fixed when mpv exits before exit-listener registration.

Risk/impact context:

- mpv plugin loading path changed from single-file to module directory; packaging/install paths must stay consistent with release assets.
- Async control/AniSkip path changes reduce blocking but can surface timing differences; regression checks added for cold start, file-load gating, and explicit refresh behavior.
<!-- SECTION:DESCRIPTION:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

AniSkip gate/async update delivered in plugin runtime:

- `plugin/subminer/lifecycle.lua`: deferred AniSkip fetch and overlay-start trigger.
- `plugin/subminer/aniskip.lua`: async lookup pipeline + context guard + session caches.
- `plugin/subminer/environment.lua`: async app-running detection with short cache.
- `plugin/subminer/messages.lua`: explicit script-message trigger wiring.

Regression coverage updated:

- `scripts/test-plugin-start-gate.lua` now verifies:
  - no sync `ps`/`curl` on non-context file load
  - no AniSkip network lookup on non-context file load
  - script-message refresh forces async AniSkip lookup

Validation run:

- `bun run test:plugin:src` pass.
<!-- SECTION:FINAL_SUMMARY:END -->
