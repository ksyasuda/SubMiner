---
id: TASK-73
title: 'MPV plugin: split into modules and optimize startup/command runtime'
status: In Progress
assignee: []
created_date: '2026-02-28 20:50'
labels: []
dependencies: []
references:
  - plugin/subminer/main.lua
  - plugin/subminer/bootstrap.lua
  - plugin/subminer/process.lua
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
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Scope: Replace monolithic `plugin/subminer.lua` with modular plugin runtime; optimize command execution paths; align install/docs/tests; fix launcher smoke instability.

Delivered behavior:
- Full plugin cutover to `plugin/subminer/main.lua` + module directory (no runtime compatibility shim with old monolith file).
- Process/control command path moved toward async subprocess usage for non-start actions (`stop`, `toggle`, `settings`, restart stop leg), reducing synchronous blocking in mpv script runtime.
- Startup/runtime loading updated with lazy module initialization via bootstrap proxies.
- Plugin install flow updated to copy `plugin/subminer/` directory and remove legacy `~/.config/mpv/scripts/subminer.lua` file.
- Added plugin gate script wiring to package scripts (`test:plugin:src`) and launcher test flow.
- Smoke tests stabilized across sandbox environments where UNIX socket bind can return `EPERM` while preserving normal-path assertions.
- Playback command cleanup race fixed when mpv exits before exit-listener registration.

Risk/impact context:
- mpv plugin loading path changed from single-file to module directory; packaging/install paths must stay consistent with release assets.
- Async control path changes reduce blocking but can surface timing differences; regression checks added for cold start and launcher smoke behavior.
<!-- SECTION:DESCRIPTION:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Startup fixed-sleep removal delivered in launcher + plugin start paths:
- `launcher/mpv.ts:startOverlay` no longer blocks on hard `2000ms`; now uses bounded readiness/event checks:
  - bounded wait for mpv IPC socket readiness
  - bounded wait for spawned overlay command settle (`exit`/`error`)
- `plugin/subminer/process.lua` removed fixed `0.35s` texthooker delay and fixed `0.6s` startup visibility delay.
- Overlay start now uses short bounded retries on non-ready start failures, then applies startup overlay visibility once start succeeds.

Regression coverage added:
- `launcher/mpv.test.ts` adds guard that `startOverlay` resolves quickly when readiness arrives.
- `scripts/test-plugin-process-start-retries.lua` asserts retry behavior and guards against reintroducing fixed `0.35/0.6` timers.

Validation run:
- `bun test launcher/mpv.test.ts` pass.
- `lua scripts/test-plugin-process-start-retries.lua` pass.
- Existing broader plugin gate suite still has unrelated failure in current tree (`scripts/test-plugin-start-gate.lua` AniSkip async curl assertion).
<!-- SECTION:FINAL_SUMMARY:END -->
