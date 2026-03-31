---
id: TASK-206
title: 'Assess latest PR #19 CodeRabbit review comments'
status: Done
assignee:
  - '@codex'
created_date: '2026-03-20 02:51'
updated_date: '2026-03-23 03:22'
labels:
  - pr-review
  - launcher
  - anki-integration
  - docs
milestone: m-1
dependencies: []
references:
  - launcher/commands/command-modules.test.ts
  - launcher/commands/stats-command.ts
  - launcher/config/cli-parser-builder.ts
  - launcher/mpv.ts
  - README.md
  - src/anki-integration.ts
  - src/anki-integration/known-word-cache.ts
priority: medium
ordinal: 125500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Validate the latest 2026-03-20 CodeRabbit review round on PR #19 against the current branch, implement only the confirmed fixes, and record which bot suggestions are stale, incorrect, or incomplete.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Each latest-round 2026-03-20 CodeRabbit inline comment on PR #19 is validated against current branch behavior and classified as actionable or not warranted
- [x] #2 Confirmed correctness issues in launcher, Anki integration, and docs are fixed with focused regression coverage where practical
- [x] #3 Targeted verification runs for the touched areas succeed or remaining unrelated failures are documented in task notes
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Pull the 2026-03-20 CodeRabbit review threads from PR #19 and validate each comment against the current branch, separating real issues from stale or incomplete bot guidance.
2. For each confirmed behavior bug, add or extend a focused failing test before changing production code; keep docs-only fixes scoped to the exact markdownlint/install issue.
3. Patch the smallest safe fixes in launcher, README, and Anki integration code, taking care not to overwrite unrelated local edits.
4. Run targeted tests and relevant SubMiner verification lanes for touched files, then record accepted versus rejected review comments in task notes and summary.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Validated the 2026-03-20 CodeRabbit PR #19 round as eight actionable items: one launcher test-name mismatch, three launcher behavior/test fixes, two README markdown/install fixes, one dead-code cleanup in Anki integration, and one real known-word cache deck-scoping bug.

Known-word cache review comment was correct in substance but needed a branch-specific fix: preserve deck->field scoping by querying per deck and carrying the allowed field list per note, rather than changing `notesInfo` shape.

Verification passed for targeted tests plus verifier docs/launcher-plugin lanes. Core verifier failed on unrelated pre-existing typecheck worktree state in `src/anki-integration/anki-connect-proxy.test.ts` (`TS2349` at line 395, `releaseProcessing?.()`), which is outside this task's touched files.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Assessed the latest 2026-03-20 CodeRabbit review round on PR #19 and applied all eight confirmed action items. Launcher behavior now surfaces non-zero stats-process exits after the startup handshake, rejects cleanup-only stats flags unless `cleanup` is selected, preserves empty quoted `mpv` args, and has updated regression coverage for each case. The known-word cache now preserves deck-specific field mappings during refresh by querying configured decks separately and extracting only the fields assigned to each deck; the unused `getPreferredWordValue` wrapper in `src/anki-integration.ts` was removed.

Documentation/test hygiene fixes also landed: the README platform badge no longer has an empty link target, Linux AppImage install instructions create `~/.local/bin` before downloads, the stats-command timing test was renamed to match actual behavior, and `launcher/picker.test.ts` now restores `XDG_DATA_HOME` safely while forcing Linux-path expectations explicitly so the file passes on macOS hosts.

Verification run:
- `bun test launcher/commands/command-modules.test.ts`
- `bun test launcher/parse-args.test.ts`
- `bun test launcher/mpv.test.ts`
- `bun test launcher/picker.test.ts`
- `bun test src/anki-integration/known-word-cache.test.ts`
- `bash .agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh README.md launcher/commands/command-modules.test.ts launcher/commands/stats-command.ts launcher/config/cli-parser-builder.ts launcher/mpv.test.ts launcher/mpv.ts launcher/parse-args.test.ts launcher/picker.test.ts src/anki-integration.ts src/anki-integration/known-word-cache.test.ts src/anki-integration/known-word-cache.ts`
- `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane docs --lane launcher-plugin --lane core README.md launcher/commands/command-modules.test.ts launcher/commands/stats-command.ts launcher/config/cli-parser-builder.ts launcher/mpv.test.ts launcher/mpv.ts launcher/parse-args.test.ts launcher/picker.test.ts src/anki-integration.ts src/anki-integration/known-word-cache.test.ts src/anki-integration/known-word-cache.ts`

Verifier results:
- `docs` lane passed (`docs:test`, `docs:build`)
- `launcher-plugin` lane passed (`test:launcher:smoke:src`, `test:plugin:src`)
- `core/typecheck` failed on unrelated existing worktree changes in `src/anki-integration/anki-connect-proxy.test.ts(395,5)`: `TS2349 This expression is not callable. Type 'never' has no call signatures.`
- Verifier artifacts: `.tmp/skill-verification/subminer-verify-20260319-195752-RNLVgE`
<!-- SECTION:FINAL_SUMMARY:END -->
