---
id: TASK-117.1
title: Harden AI subtitle fix against non-SRT model responses
status: Done
assignee:
  - '@codex'
created_date: '2026-03-08 08:22'
updated_date: '2026-03-08 08:25'
labels: []
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/launcher/youtube/subtitle-fix-ai.ts
  - /Users/sudacode/projects/japanese/SubMiner/launcher/youtube/srt.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/launcher/youtube/subtitle-fix-ai.test.ts
parent_task_id: TASK-117
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Prevent optional YouTube AI subtitle post-processing from bailing out whenever the model returns usable cue text in a non-SRT wrapper or text-only format. The launcher should recover safe cases, preserve original timing, and fall back cleanly when the response cannot be mapped back to the source cues.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 AI subtitle fixing accepts safe AI responses that omit SRT framing but still provide one corrected text payload per original cue while preserving original cue timing.
- [x] #2 AI subtitle fixing still rejects responses that cannot be mapped back to the original cue batch without guessing and falls back to the raw subtitle file with a warning.
- [x] #3 Automated tests cover wrapped-SRT and text-only AI responses plus an unrecoverable invalid response case.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing tests in launcher/youtube/subtitle-fix-ai.test.ts for three cases: wrapped valid SRT, text-only one-block-per-cue output, and unrecoverable invalid output.
2. Extend launcher/youtube/subtitle-fix-ai.ts with a small response-normalization path that first strips markdown/code-fence wrappers, then accepts deterministic text-only cue batches only when they map 1:1 to the original cues without changing timestamps.
3. Keep existing safety rules: preserve cue count and timing, log a warning, and fall back to the raw subtitle file when normalization cannot recover a trustworthy batch.
4. Run focused launcher unit tests for subtitle-fix-ai and SRT parsing; expand only if the change affects adjacent behavior.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented deterministic AI subtitle-response recovery for fenced SRT, embedded SRT payloads, and text-only 1:1 cue batches while preserving original timing and existing fallback behavior.

Verification: bun test launcher/youtube/*.test.ts passed; bun run typecheck passed; repo-wide format check still reports unrelated pre-existing warnings in launcher/youtube/orchestrator.ts and scripts/build-changelog*.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Hardened the launcher AI subtitle-fix path so it can recover deterministic non-SRT model responses instead of immediately falling back. Added `parseAiSubtitleFixResponse` in `launcher/youtube/subtitle-fix-ai.ts` to normalize markdown-fenced or embedded SRT payloads first, then accept text-only responses only when they map 1:1 onto the original cue batch and preserve source timings. Added regression coverage in `launcher/youtube/subtitle-fix-ai.test.ts` for fenced SRT, text-only cue batches, and unrecoverable invalid output, plus a changelog fragment in `changes/task-117.1.md`.

Verification: `bun test launcher/youtube/*.test.ts`, `bun run typecheck`, `bunx prettier --check launcher/youtube/subtitle-fix-ai.ts launcher/youtube/subtitle-fix-ai.test.ts`, and `bun run changelog:lint` passed. Repo-wide `bun run format:check:src` still reports unrelated pre-existing warnings in `launcher/youtube/orchestrator.ts` and `scripts/build-changelog*`.
<!-- SECTION:FINAL_SUMMARY:END -->
