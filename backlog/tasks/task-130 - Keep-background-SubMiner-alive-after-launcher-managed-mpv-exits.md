---
id: TASK-130
title: Keep background SubMiner alive after launcher-managed mpv exits
status: Done
assignee:
  - codex
created_date: '2026-03-08 10:08'
updated_date: '2026-03-08 11:00'
labels:
  - bug
  - launcher
  - mpv
  - overlay
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

The launcher currently tears down the running SubMiner background process when a launcher-managed mpv session exits. Background SubMiner should remain alive so a later mpv instance can reconnect and request the overlay without restarting the app.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Closing a launcher-managed mpv session does not send `--stop` to the running SubMiner background process.
- [x] #2 Closing a launcher-managed mpv session does not SIGTERM the tracked SubMiner process just because mpv exited.
- [x] #3 Launcher cleanup still terminates mpv and launcher-owned helper children without regressing existing overlay start behavior.
- [x] #4 Automated tests cover the no-stop-on-mpv-exit behavior.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Add a launcher regression test that proves mpv exit no longer triggers SubMiner `--stop` or launcher SIGTERM of the tracked overlay process.
2. Update launcher teardown so normal mpv-session cleanup only stops mpv/helper children and preserves the background SubMiner process for future reconnects.
3. Run the focused launcher tests and smoke coverage for the affected behavior, then record results in the task.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Split launcher cleanup so normal mpv-session shutdown no longer sends `--stop` to SubMiner or SIGTERM to the tracked overlay process. Added `cleanupPlaybackSession()` for mpv/helper-child cleanup only, and switched playback finalization to use it.

Updated launcher smoke coverage to assert the background app stays alive after mpv exits, and added a focused unit regression for the new cleanup path.

Validation: `bun test launcher/mpv.test.ts launcher/smoke.e2e.test.ts` passed; `bun run typecheck` passed. `bun run test:launcher:unit:src` still reports an unrelated pre-existing failure in `launcher/aniskip-metadata.test.ts`.

Added changelog fragment `changes/task-130.md` for the launcher fix and verified it with `bun run changelog:lint`.

User verified the bug still reproduces when closing playback with `q`. Root cause narrowed further: the mpv plugin `plugin/subminer/lifecycle.lua` calls `process.stop_overlay()` on mpv `shutdown`, which still sends SubMiner `--stop` even after launcher cleanup was fixed.

Patched the remaining stop path in `plugin/subminer/lifecycle.lua`: mpv `shutdown` no longer calls `process.stop_overlay()`. Pressing mpv `q` should now preserve the background app and only tear down the mpv session.

Validation update: `lua scripts/test-plugin-start-gate.lua` passed after adding a shutdown regression, and `bun test launcher/mpv.test.ts launcher/smoke.e2e.test.ts` still passed.

Fixed a second-instance reconnect bug in `src/core/services/cli-command.ts`: `--start` on an already-initialized running instance now still updates the MPV socket path and reconnects the MPV client instead of treating the command as a no-op. This keeps the already-warmed background app reusable for later mpv launches.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Kept the background SubMiner process reusable across both mpv shutdown and later reconnects. The first fix separated launcher playback cleanup from full app shutdown. The second fix removed the mpv plugin `shutdown` stop call so default mpv `q` no longer sends SubMiner `--stop`. The third fix corrected second-instance CLI handling so `--start` on an already-running, already-initialized instance still reconnects MPV instead of being ignored.

Net effect: background SubMiner can stay alive, keep its warm state, and reconnect to later mpv instances without rerunning startup/warmup work in a fresh app instance.

Coverage now includes: launcher playback cleanup (`launcher/mpv.test.ts`), launcher smoke reconnect/keep-alive flow (`launcher/smoke.e2e.test.ts`), mpv plugin shutdown preservation (`scripts/test-plugin-start-gate.lua`), and second-instance start/reconnect behavior (`src/core/services/cli-command.test.ts`).

Tests run:

- `bun test src/core/services/cli-command.test.ts launcher/mpv.test.ts launcher/smoke.e2e.test.ts`
- `lua scripts/test-plugin-start-gate.lua`
- `bun run typecheck`
- `bun run changelog:lint`

Note: the broader `bun run test:launcher:unit:src` lane still has an unrelated pre-existing failure in `launcher/aniskip-metadata.test.ts`.

<!-- SECTION:FINAL_SUMMARY:END -->
