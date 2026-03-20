---
id: TASK-85
title: >-
  Add launcher dictionary subcommand and initial AniList character dictionary
  zip generation
status: Done
assignee: []
created_date: '2026-03-03 08:47'
updated_date: '2026-03-16 05:13'
labels: []
dependencies: []
priority: high
ordinal: 96500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Implement initial character dictionary flow: launcher `dictionary` subcommand, app `--dictionary` command, AniList media resolution from current playback, Yomitan zip generation to local file, and local cache to avoid repeated API fetches for same AniList id. Manual Yomitan import path only in this phase.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launcher supports `dictionary` (and alias) and forwards to app command path.
- [x] #2 App CLI accepts `--dictionary` and dispatches to dictionary runtime command.
- [x] #3 Dictionary command resolves current anime to AniList id, generates Yomitan-compatible zip, and logs output path for manual load.
- [x] #4 Generated dictionaries are cached by AniList id so repeated commands reuse existing zip when available.
- [x] #5 Backlog task is updated with implementation notes and completion summary for this phase.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented launcher `dictionary`/`dict` subcommand parsing and normalized args flow (`launcher/config/cli-parser-builder.ts`, `launcher/config/args-normalizer.ts`, `launcher/types.ts`).

Added launcher command dispatch (`launcher/commands/dictionary-command.ts`) and wired `launcher/main.ts` to forward `--dictionary` (plus non-default `--log-level`) to app binary.

Added app CLI flag `--dictionary` in parser/help and startup routing (`src/cli/args.ts`, `src/cli/help.ts`).

Added dictionary runtime service (`src/main/character-dictionary-runtime.ts`) that resolves AniList media id from current playback guess, fetches AniList character edges, builds Yomitan-compatible banks/index, writes zip, and caches by AniList id in user data.

Threaded dictionary generation dependency through CLI runtime/context builders and `src/main.ts` context composition so command executes from launcher/app entrypoints.

Added/updated tests for parser, command modules, launcher main forwarding, CLI command dispatch, and context/deps wiring. Updated docs for launcher/usage command lists to include dictionary subcommand.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Initial phase shipped: `subminer dictionary` now routes to `SubMiner.AppImage --dictionary`, generates a Yomitan-importable character dictionary zip for the current anime (AniList-based), logs zip output path for manual import, and reuses cached zips by AniList id to avoid repeated API fetches.
<!-- SECTION:FINAL_SUMMARY:END -->
