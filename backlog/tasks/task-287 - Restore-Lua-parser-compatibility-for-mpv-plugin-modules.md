---
id: TASK-287
title: Restore Lua parser compatibility for mpv plugin modules
status: Done
assignee: []
created_date: '2026-04-11 21:25'
updated_date: '2026-04-11 21:29'
labels:
  - bug
  - mpv-plugin
  - lua
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Users with Lua runtimes that do not accept the current `goto continue` pattern in the mpv plugin should be able to load the plugin without syntax errors. Remove the parser-incompatible control-flow usage from the affected plugin modules without changing plugin behavior.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 The mpv plugin source no longer relies on parser-incompatible `goto continue` labels in the affected Lua modules.
- [x] #2 Automated coverage fails on the old parser-incompatible source and passes once the compatibility fix is in place.
- [x] #3 Existing plugin start/gate verification still passes after the compatibility fix.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Reused existing local cleanups in `plugin/subminer/hover.lua` and `plugin/subminer/environment.lua` to remove `goto continue` / `::continue::` control flow without behavior changes.

Added `scripts/test-plugin-lua-compat.lua` and wired it into `test:plugin:src`; the regression checks reject the legacy pattern structurally and verify parse success with `luajit` when available.

Verification run on 2026-04-11: `lua scripts/test-plugin-lua-compat.lua` ✅, `bun run test:plugin:src` ✅, `bun run changelog:lint` ✅, `bun run typecheck` ✅, `bun run test:env` ✅, `bun run build` ✅, `bun run test:smoke:dist` ✅.

`bun run test:fast` remains red for unrelated existing immersion-tracker assertions in `src/core/services/immersion-tracker/__tests__/query-split-modules.test.ts` and `src/core/services/immersion-tracker/__tests__/query.test.ts` (`tsMs`/`lastWatchedMs` observed as `-2147483648`).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Removed parser-incompatible `goto continue` usage from the affected mpv Lua plugin modules, added a dedicated Lua compatibility regression script to the plugin test lane, and added a changelog fragment for the user-visible fix. Requested plugin verification is green; unrelated existing `test:fast` immersion-tracker failures remain outside this task.
<!-- SECTION:FINAL_SUMMARY:END -->
