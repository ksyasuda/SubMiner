---
id: TASK-208
title: 'Assess newest PR #19 CodeRabbit round after 1227706'
status: Done
assignee:
  - '@codex'
created_date: '2026-03-20 03:37'
updated_date: '2026-03-20 03:47'
labels:
  - pr-review
  - launcher
  - anki-integration
milestone: m-1
dependencies: []
references:
  - launcher/commands/stats-command.ts
  - launcher/mpv.ts
  - src/anki-integration.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Validate the newest 2026-03-20 03:23 CodeRabbit review round on PR #19 after commit `1227706`, implement only the confirmed fixes, and record any bot suggestions that are stale or technically incomplete.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Each newest-round CodeRabbit inline comment posted after commit `1227706` is validated against current branch behavior and classified as actionable or not warranted
- [x] #2 Confirmed issues are fixed with focused regression coverage where practical
- [x] #3 Targeted verification runs for the touched areas succeed or remaining unrelated failures are documented
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Pull the three newest CodeRabbit inline threads posted after commit `1227706` and restate each finding against the current branch code.
2. For each confirmed behavior bug, add or extend a focused failing test before changing production code; reject any stale or incorrect bot suggestion with notes.
3. Patch the smallest safe fixes in `launcher/commands/stats-command.ts`, `launcher/mpv.ts`, and/or `src/anki-integration.ts` as warranted, without disturbing unrelated local edits.
4. Run targeted tests and the cheapest sufficient verifier lanes, then record accepted versus rejected comments in task notes and summary.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Validated the newest 2026-03-20 03:23 CodeRabbit round as three comments: two actionable launcher issues and one non-warranted Anki suggestion.

Accepted fixes: cancel the pending stats response poll when the attached app exits non-zero before startup response, and surface `spawnSync()` launch/stop errors in launcher mpv helpers instead of treating `result.status ?? 0` / ignored status as success.

Rejected fix: the `src/anki-integration.ts` / card-creation suggestion would double count locally mined cards. Local sentence mining already records stats in `src/main/runtime/anki-actions.ts` when `mineSentenceCardCore` returns `true`; adding a second callback in card creation would increment tracker counts twice for the same card.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Assessed the newest CodeRabbit PR #19 round after commit `1227706` and fixed the two confirmed launcher regressions. `runStatsCommand()` now gives the startup response waiter an abort signal and cancels the polling loop immediately when the attached app exits non-zero before startup response, covering both the normal stats startup race and the cleanup/startup race. `launchTexthookerOnly()` now fails non-zero when `spawnSync()` reports an execution error, and `stopOverlay()` logs a warning when the stop command cannot be spawned or exits non-zero instead of silently treating that path as success.

One bot comment was intentionally rejected: recording mined-card stats inside the direct card-creation path would double count locally mined cards, because the successful local mining flow already records cards in `src/main/runtime/anki-actions.ts` after `mineSentenceCardCore()` returns `true`.

Verification run:
- `bun test launcher/commands/command-modules.test.ts`
- `bun test launcher/mpv.test.ts`
- `bun run typecheck`
- `bash .agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh launcher/commands/stats-command.ts launcher/commands/command-modules.test.ts launcher/mpv.ts launcher/mpv.test.ts`
- `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane launcher-plugin launcher/commands/stats-command.ts launcher/commands/command-modules.test.ts launcher/mpv.ts launcher/mpv.test.ts`

Verifier result:
- `launcher-plugin` lane passed (`test:launcher:smoke:src`, `test:plugin:src`)
- `typecheck` passed
- Verifier artifacts: `.tmp/skill-verification/subminer-verify-20260319-204639-dzUj16`
<!-- SECTION:FINAL_SUMMARY:END -->
