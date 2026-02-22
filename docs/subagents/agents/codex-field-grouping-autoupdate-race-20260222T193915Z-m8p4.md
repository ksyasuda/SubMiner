# codex-field-grouping-autoupdate-race-20260222T193915Z-m8p4

- alias: `codex-field-grouping-autoupdate-race`
- mission: Fix Kiku field-grouping merge race with auto-update so note enrichment completes before duplicate merge.
- status: `handoff`
- last_update_utc: `2026-02-22T19:41:20Z`
- backlog_ticket: `TASK-94` (`backlog/tasks/task-94 - Fix-Kiku-duplicate-detection-for-Yomitan-marked-duplicates.md`) for duplicate/grouping workflow scope.

## Intent

- Reproduce reported ordering bug in `NoteUpdateWorkflow`.
- Add regression test first (TDD): update fields before field-grouping merge.
- Patch workflow ordering with minimal behavior change.

## Planned Files

- `src/anki-integration/note-update-workflow.test.ts`
- `src/anki-integration/note-update-workflow.ts`

## Assumptions

- Bug path: auto-update flow enters duplicate merge before sentence/audio/image enrichment completes.
- Existing field-grouping handlers remain source of merge/delete behavior; only invocation timing changes.

## Progress Log

- 2026-02-22T19:39:15Z: session start; read subagent index/collaboration; identified target workflow and tests.
- 2026-02-22T19:40:15Z: added failing TDD regression in `src/anki-integration/note-update-workflow.test.ts` asserting update-before-auto-merge order.
- 2026-02-22T19:40:45Z: patched `src/anki-integration/note-update-workflow.ts` to defer field-grouping execution until after enrichment/update; refresh note info before merge.
- 2026-02-22T19:41:20Z: verification green:
  - `bun test src/anki-integration/note-update-workflow.test.ts`
  - `bun test src/anki-integration/field-grouping-workflow.test.ts src/anki-integration.test.ts`

## Files Touched

- `src/anki-integration/note-update-workflow.ts`
- `src/anki-integration/note-update-workflow.test.ts`
- `docs/subagents/INDEX.md`
- `docs/subagents/agents/codex-field-grouping-autoupdate-race-20260222T193915Z-m8p4.md`

## Decisions

- Keep duplicate detection early (single lookup), but run merge/manual handlers only after card update path completes.
- Refresh `noteInfo` after update so grouping handlers operate on enriched fields.

## Blockers

- None.

## Next Step

- Optional: add an integration-level test for keep-both (`kikuDeleteDuplicateInAuto=false`) flow to assert both notes retain enriched data.
