---
id: TASK-124
title: >-
  Remove YouTube subtitle generation modes and make YouTube playback always
  generate/load subtitles
status: Done
assignee:
  - codex
created_date: '2026-03-08 07:18'
updated_date: '2026-03-16 05:13'
labels:
  - launcher
  - youtube
  - subtitles
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/launcher/commands/playback-command.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/launcher/config/args-normalizer.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/launcher/config/youtube-subgen-config.ts
  - /Users/sudacode/projects/japanese/SubMiner/launcher/types.ts
  - /Users/sudacode/projects/japanese/SubMiner/config.example.jsonc
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/config/definitions/options-integrations.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/config/resolve/subtitle-domains.ts
priority: high
ordinal: 61500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Simplify launcher YouTube playback by removing the configurable subtitle generation mode. For YouTube targets, the launcher should treat subtitle generation/loading as the canonical behavior instead of supporting off/preprocess/automatic branches. This change should remove the unreliable automatic/background path and the mode concept from config/CLI/env/docs, while preserving the core YouTube subtitle generation pipeline and mpv loading flow.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launcher playback no longer supports or branches on a YouTube subtitle generation mode; YouTube URLs follow a single generation-and-load flow.
- [x] #2 Configuration, CLI parsing, and environment handling no longer expose a YouTube subtitle generation mode option, and stale automatic/preprocess/off values are not part of the supported interface.
- [x] #3 Tests cover the new single-flow behavior and the removal of mode parsing/branching.
- [x] #4 User-facing config/docs/examples are updated to reflect the removed mode concept.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Remove the YouTube subtitle generation mode concept from launcher/shared types, config parsing, CLI options, and environment normalization so no supported interface accepts automatic/preprocess/off.
2. Update playback orchestration so YouTube targets always run subtitle generation/loading before mpv startup and delete the background automatic path.
3. Adjust mpv YouTube URL argument construction to no longer branch on mode while preserving subtitle/audio language behavior and preloaded subtitle file injection.
4. Add/modify tests first to cover removed mode parsing and the single YouTube preload flow, then update config/docs/examples to match the simplified interface.
5. Run focused launcher/config tests plus typecheck, then summarize any remaining gaps.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Removed launcher/shared youtubeSubgen.mode handling and collapsed YouTube playback onto a single preload-before-mpv subtitle generation flow.

Added launcher integration coverage proving YouTube subtitle generation runs before mpv startup and that the removed --mode flag now errors.

Verification: bun test launcher/config-domain-parsers.test.ts launcher/parse-args.test.ts launcher/mpv.test.ts launcher/main.test.ts src/config/config.test.ts; bun run test:config:src; bun run typecheck.

Broader repo checks still show pre-existing issues outside this change: bun run test:launcher:unit:src fails in launcher/aniskip-metadata.test.ts (MAL id assertion), and format scope check reports unrelated existing files launcher/youtube/orchestrator.ts, scripts/build-changelog.test.ts, scripts/build-changelog.ts.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Removed the launcher YouTube subtitle generation mode surface so YouTube playback now always runs the subtitle generation pipeline before starting mpv. The launcher no longer accepts youtubeSubgen.mode from shared config, CLI, or env normalization, and the old automatic/background loading path has been deleted from playback.

Updated mpv YouTube startup options to keep manual subtitle discovery enabled without requesting auto subtitles, and refreshed user-facing config/docs to describe a single YouTube subtitle generation flow. Added regression coverage for mode removal, config/template cleanup, and launcher ordering so YouTube subtitle work is confirmed to happen before mpv launch.

Verification: bun test launcher/config-domain-parsers.test.ts launcher/parse-args.test.ts launcher/mpv.test.ts launcher/main.test.ts src/config/config.test.ts; bun run test:config:src; bun run typecheck. Broader unrelated repo issues remain in launcher/aniskip-metadata.test.ts and existing formatting drift in launcher/youtube/orchestrator.ts plus scripts/build-changelog files.
<!-- SECTION:FINAL_SUMMARY:END -->
