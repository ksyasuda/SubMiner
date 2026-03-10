---
id: TASK-127
title: Skip AniSkip lookup for YouTube and URL playback targets
status: Done
assignee:
  - '@codex'
created_date: '2026-03-08 08:24'
updated_date: '2026-03-08 10:12'
labels:
  - bug
  - launcher
  - youtube
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/launcher/mpv.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/launcher/commands/playback-command.ts
  - /Users/sudacode/projects/japanese/SubMiner/launcher/mpv.test.ts
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Prevent launcher playback from attempting AniSkip metadata resolution when the user is playing a YouTube target or any URL target. AniSkip only works for local anime files, so URL-driven playback and YouTube subtitle-generation flows should bypass it entirely.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Launcher playback skips AniSkip metadata resolution for explicit URL targets, including YouTube URLs.
- [x] #2 YouTube subtitle-generation playback does not invoke AniSkip lookup before mpv launch.
- [x] #3 Automated launcher tests cover the URL/YouTube skip behavior.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Add a launcher mpv unit test that intercepts AniSkip resolution and proves URL/YouTube playback does not call it before spawning mpv.
2. Run the focused launcher mpv test to confirm the new case fails or exposes the current gap.
3. Patch launcher playback/AniSkip gating so URL and YouTube subtitle-generation paths always bypass AniSkip lookup.
4. Re-run focused launcher tests and record the verification results in task notes.

5. Add a Lua plugin regression test covering overlay-start on URL playback so AniSkip never runs after auto-start.

6. Patch plugin/subminer/aniskip.lua to short-circuit all AniSkip lookup triggers for remote URL media paths.

7. Re-run plugin regression plus touched launcher checks and update the task summary with the plugin-side fix.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Added explicit AniSkip gating in launcher/mpv.ts via shouldResolveAniSkipMetadata(target, targetKind, preloadedSubtitles).

URL targets now always bypass AniSkip. File targets with preloaded subtitles also bypass AniSkip, covering YouTube subtitle-preload playback.

Added launcher/mpv.test.ts coverage for local-file vs URL vs preloaded-subtitle AniSkip gating.

Verification: bun test launcher/mpv.test.ts passed.

Verification: bun run typecheck passed.

Verification: bunx prettier --check launcher/mpv.ts launcher/mpv.test.ts passed.

Verification: bun run changelog:lint passed.

Verification: bun run test:launcher:unit:src remains blocked by unrelated existing failure in launcher/aniskip-metadata.test.ts (`resolveAniSkipMetadataForFile resolves MAL id and intro payload`: expected malId 1234, got null).

Added plugin regression in scripts/test-plugin-start-gate.lua for URL playback with auto-start/overlay-start; it now asserts no MAL or AniSkip curl requests occur.

Patched plugin/subminer/aniskip.lua to short-circuit AniSkip lookup for remote media paths (`scheme://...`), which covers YouTube URL playback inside the mpv plugin lifecycle.

Verification: lua scripts/test-plugin-start-gate.lua passed.

Verification: bun run test:plugin:src passed.

Verification: bun test launcher/mpv.test.ts passed after plugin-side fix.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Fixed AniSkip suppression end-to-end for URL playback. The launcher now skips AniSkip before mpv launch, and the mpv plugin now also refuses AniSkip lookups for remote URL media during file-loaded, overlay-start, or later refresh triggers. Added regression coverage in both launcher/mpv.test.ts and scripts/test-plugin-start-gate.lua, plus a changelog fragment. Wider `bun run test:launcher:unit:src` is still blocked by the unrelated existing launcher/aniskip-metadata.test.ts MAL-id failure.

<!-- SECTION:FINAL_SUMMARY:END -->
