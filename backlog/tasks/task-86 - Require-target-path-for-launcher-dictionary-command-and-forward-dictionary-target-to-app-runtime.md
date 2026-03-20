---
id: TASK-86
title: >-
  Require target path for launcher dictionary command and forward dictionary
  target to app runtime
status: Done
assignee: []
created_date: '2026-03-03 09:22'
updated_date: '2026-03-16 05:13'
labels: []
dependencies: []
priority: high
ordinal: 95500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Change dictionary flow so launcher uses `subminer dictionary <file-or-directory>` and forwards target to app without playback launch. Keep direct app `--dictionary` behavior for in-session/mpv-triggered use, adding optional `--dictionary-target` path override.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launcher `dictionary`/`dict` requires a target path argument and parses optional log level.
- [x] #2 Launcher forwards target to app as `--dictionary-target <path>` together with `--dictionary`.
- [x] #3 App CLI parses optional `--dictionary-target` and dictionary command passes it into dictionary runtime.
- [x] #4 Dictionary runtime resolves explicit file/directory target for AniList guess flow; falls back to current playback when target is absent.
- [x] #5 Docs/tests are updated for new launcher invocation syntax.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Launcher dictionary subcommand now requires a positional target path (`subminer dictionary <path>` / `dict <path>`) via commander argument wiring in `launcher/config/cli-parser-builder.ts`.

Added `dictionaryTarget` flow in launcher normalization/types and path validation (must be existing local file or directory) in `launcher/config/args-normalizer.ts`.

Launcher dictionary command now forwards `--dictionary --dictionary-target <resolved-path>` to app binary in `launcher/commands/dictionary-command.ts`.

App CLI now parses optional `--dictionary-target` in `src/cli/args.ts`; dictionary handler passes target through to runtime in `src/core/services/cli-command.ts`.

Dictionary runtime now resolves explicit target inputs in `src/main/character-dictionary-runtime.ts`: file target uses file path/title; directory target selects first video file recursively (fallback to directory name), and still falls back to current playback when target is absent.

Updated context/dependency pass-through for dictionary target argument (`src/main.ts`, `src/main/runtime/cli-command-context-main-deps.ts`).

Updated tests/docs for new syntax and forwarding behavior (`launcher/main.test.ts`, `launcher/parse-args.test.ts`, `launcher/commands/command-modules.test.ts`, `src/cli/args.test.ts`, `src/cli/help.test.ts`, `docs/usage.md`, `docs/launcher-script.md`).

Follow-up noise reduction: dictionary commands now opt into lightweight startup path by extending `shouldSkipHeavyStartup` in `src/main.ts` to include `initialArgs.dictionary`. This skips heavy app-ready initialization (mpv client creation/background warmups/overlay bootstrap) for dictionary CLI runs.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Launcher dictionary flow now uses explicit targets: run `subminer dictionary <file-or-directory>`. It forwards target to app and performs dictionary generation without depending on currently playing media. Direct app `--dictionary` remains available for in-session/mpv-triggered workflows, with optional `--dictionary-target` override support.
<!-- SECTION:FINAL_SUMMARY:END -->
