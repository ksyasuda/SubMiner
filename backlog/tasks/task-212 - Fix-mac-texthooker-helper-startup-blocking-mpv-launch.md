---
id: TASK-212
title: Fix mac texthooker helper startup blocking mpv launch
status: Done
assignee: []
created_date: '2026-03-20 08:27'
updated_date: '2026-03-23 03:22'
labels:
  - bug
  - macos
  - startup
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/src/core/services/startup.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/main.ts
  - /Users/sudacode/projects/japanese/SubMiner/plugin/subminer/process.lua
priority: high
ordinal: 140500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`subminer` mpv auto-start on mac can stall before the video is usable because the helper process launched with `--texthooker` still runs heavy app-ready startup. Recent logs show the helper loading the Yomitan Chromium extension, emitting `Permission 'contextMenus' is unknown` warnings, then hitting Chromium runtime errors before SubMiner signals readiness back to the mpv plugin. The texthooker helper should take the minimal startup path needed to serve texthooker traffic without loading overlay/window-only startup work that can crash or delay readiness.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launching SubMiner with `--texthooker` avoids heavy app-ready startup work that is not required for texthooker helper mode.
- [x] #2 A regression test covers texthooker helper startup so it fails if Yomitan extension loading is reintroduced on that path.
- [x] #3 The change preserves existing startup behavior for non-texthooker app launches.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Follow-up: user confirmed the root issue is the plugin auto-start ordering. Adjust mpv plugin sequencing so `--start` launches before any separate `--texthooker` helper, then verify plugin regressions still pass.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed the mac mpv startup hang caused by the `--texthooker` helper taking the full app-ready path. `runAppReadyRuntime` now fast-paths texthooker-only mode through minimal startup (`reloadConfig` plus CLI handling) so it no longer loads Yomitan or first-run setup work before serving texthooker traffic. Added regression coverage in `src/core/services/app-ready.test.ts`, then verified with `bun test src/core/services/app-ready.test.ts src/core/services/startup.test.ts`, `bun test src/cli/args.test.ts src/main/early-single-instance.test.ts src/main/runtime/stats-cli-command.test.ts`, and `bun run typecheck`.
<!-- SECTION:FINAL_SUMMARY:END -->
