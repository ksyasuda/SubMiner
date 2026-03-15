---
id: TASK-174
title: Fix missing frequency highlights for merged tokenizer tokens
status: In Progress
assignee:
  - codex
created_date: '2026-03-15 10:18'
updated_date: '2026-03-15 10:40'
labels:
  - bug
  - tokenizer
  - frequency-highlighting
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/src/core/services/tokenizer.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/tokenizer/parser-selection-stage.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/tokenizer/yomitan-parser-runtime.ts
  - /Users/sudacode/projects/japanese/SubMiner/scripts/get_frequency.ts
  - /Users/sudacode/projects/japanese/SubMiner/scripts/test-yomitan-parser.ts
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Frequency highlighting can miss words that should color within the configured top-X limit when tokenizer candidate selection keeps merged Yomitan units that combine a content word with trailing function text. The annotation stage then conservatively clears frequency for the whole merged token, so visible high-frequency words lose highlighting. The standalone debug CLIs are also failing to initialize the shared Yomitan runtime, which blocks reliable repro for this class of bug.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Tokenizer no longer drops frequency highlighting for content words in merged-token cases where a better scanning parse candidate would preserve highlightable tokens.
- [ ] #2 A regression test covers the reported sentence shape and fails before the fix.
- [ ] #3 The standalone frequency/parser debug path can initialize the shared Yomitan runtime well enough to reproduce tokenizer output instead of immediately reporting runtime/session wiring errors.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a regression test for the reported merged-token frequency miss, centered on Yomitan scanning candidate selection and downstream frequency annotation.
2. Update tokenizer candidate selection so merged content+function tokens do not win over candidates that preserve highlightable content tokens.
3. Repair the standalone frequency/parser debug scripts so their Electron/Yomitan runtime wiring matches current shared runtime expectations.
4. Verify with targeted tokenizer/parser tests and the standalone debug repro command.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Initial triage: shared frequency class logic looks correct; likely failure is upstream tokenizer candidate selection producing merged content+function tokens that annotation later excludes from frequency. Standalone debug scripts also fail to initialize a usable Electron/Yomitan runtime, blocking reliable repro from the current CLI path.

Repro after fixing the standalone Electron wrapper does not support the original highlight claim for `誰でもいいから かかってこいよ`: the tokenizer reports `かかってこい` with `frequencyRank` 63098, so it correctly stays uncolored at `--color-top-x 10000` and becomes colorable once the threshold is raised above that rank. The concrete bug fixed in this pass is the standalone Electron debug path: package scripts now unset `ELECTRON_RUN_AS_NODE`, and the scripts normalize Electron imports/guards so `get-frequency:electron` can reach real Electron/Yomitan runtime state instead of immediately falling back to Node-mode diagnostics. `test-yomitan-parser:electron` still shows extension/service-worker issues against the existing profile and was not stabilized in this pass.
<!-- SECTION:NOTES:END -->
