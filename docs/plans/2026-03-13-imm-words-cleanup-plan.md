# Imm Words Cleanup Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Fix `imm_words` so only allowed vocabulary tokens are persisted with POS metadata, and add an on-demand stats cleanup command that removes existing bad vocabulary rows.

**Architecture:** Thread processed `SubtitleData.tokens` into immersion tracking instead of extracting vocabulary from raw subtitle text. Store POS metadata on `imm_words`, reuse the existing POS exclusion logic for persistence and cleanup, and expose cleanup through the existing launcher/app stats command surface.

**Tech Stack:** TypeScript, Bun, SQLite/libsql wrapper, Commander-based launcher CLI

---

### Task 1: Lock the broken behavior down with failing tests

**Files:**
- Modify: `src/core/services/immersion-tracker-service.test.ts`
- Modify: `src/core/services/immersion-tracker/__tests__/query.test.ts`
- Modify: `launcher/parse-args.test.ts`
- Modify: `src/main/runtime/stats-cli-command.test.ts`

**Steps:**
1. Add a tracker regression test that records subtitle tokens with mixed POS and asserts excluded tokens are not written while allowed tokens retain POS metadata.
2. Add a cleanup/query test that seeds valid and invalid `imm_words` rows and asserts vocab cleanup deletes only invalid rows.
3. Add launcher parse tests for `subminer stats cleanup`, `subminer stats cleanup -v`, and default vocab cleanup mode.
4. Add app-side stats CLI tests for dispatching cleanup vs dashboard launch.

### Task 2: Fix live vocabulary persistence

**Files:**
- Modify: `src/core/services/immersion-tracker-service.ts`
- Modify: `src/core/services/immersion-tracker/types.ts`
- Modify: `src/core/services/immersion-tracker/storage.ts`
- Modify: `src/main/runtime/mpv-main-event-main-deps.ts`
- Modify: `src/main/runtime/mpv-main-event-bindings.ts`
- Modify: `src/main/runtime/mpv-main-event-actions.ts`
- Modify: `src/main/state.ts`

**Steps:**
1. Extend immersion subtitle recording to accept processed subtitle payloads or token arrays.
2. Add `imm_words` POS columns and prepared-statement support.
3. Convert tracker inserts to use processed tokens rather than raw regex extraction.
4. Reuse existing POS/frequency exclusion rules to decide whether a token belongs in `imm_words`.

### Task 3: Add vocab cleanup service and CLI wiring

**Files:**
- Modify: `src/core/services/immersion-tracker/query.ts`
- Modify: `src/main/runtime/stats-cli-command.ts`
- Modify: `src/cli/args.ts`
- Modify: `launcher/config/cli-parser-builder.ts`
- Modify: `launcher/config/args-normalizer.ts`
- Modify: `launcher/types.ts`
- Modify: `launcher/commands/stats-command.ts`

**Steps:**
1. Add a cleanup routine that removes invalid `imm_words` rows using the same persistence filter.
2. Extend app CLI args to represent stats cleanup actions and vocab mode selection.
3. Extend launcher stats command forwarding to pass cleanup flags through the attached app flow.
4. Print a compact cleanup summary and fail cleanly on errors.

### Task 4: Verify and document the final behavior

**Files:**
- Modify: `docs-site/immersion-tracking.md`

**Steps:**
1. Update user-facing stats docs to mention `subminer stats cleanup` vocab maintenance.
2. Run the cheapest sufficient verification lanes for touched files.
3. Record exact commands/results in the task final summary before handoff.
