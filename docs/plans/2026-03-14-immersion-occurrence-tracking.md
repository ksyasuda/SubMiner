# Immersion Occurrence Tracking Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add normalized counted occurrence tracking for immersion words and kanji so stats can map each item back to anime, episode, timestamp, and subtitle line context.

**Architecture:** Introduce `imm_subtitle_lines` plus counted bridge tables from lines to canonical `imm_words` and `imm_kanji` rows. Extend the subtitle write path to persist one line row per subtitle event, retain aggregate lexeme tables, and expose reverse-mapping queries without duplicating repeated token text in storage.

**Tech Stack:** TypeScript, Bun, libsql SQLite, existing immersion tracker storage/query/service modules

---

### Task 1: Lock schema and migration shape down with failing tests

**Files:**
- Modify: `src/core/services/immersion-tracker/storage-session.test.ts`
- Modify: `src/core/services/immersion-tracker/storage.ts`
- Modify: `src/core/services/immersion-tracker/types.ts`

**Steps:**
1. Add a red test asserting `ensureSchema()` creates `imm_subtitle_lines`, `imm_word_line_occurrences`, and `imm_kanji_line_occurrences`, plus additive migration support from the previous schema version.
2. Run `bun test src/core/services/immersion-tracker/storage-session.test.ts` and confirm failure.
3. Implement the minimal schema/version/index changes.
4. Re-run the targeted test and confirm green.

### Task 2: Lock counted subtitle-line persistence down with failing tests

**Files:**
- Modify: `src/core/services/immersion-tracker-service.test.ts`
- Modify: `src/core/services/immersion-tracker-service.ts`
- Modify: `src/core/services/immersion-tracker/storage.ts`

**Steps:**
1. Add a red service test that records a subtitle line with repeated allowed words and repeated kanji, then asserts one line row plus counted bridge rows are written.
2. Run `bun test src/core/services/immersion-tracker-service.test.ts` and confirm failure.
3. Implement the minimal subtitle-line insert and counted occurrence write path.
4. Re-run the targeted test and confirm green.

### Task 3: Add reverse-mapping query tests first

**Files:**
- Modify: `src/core/services/immersion-tracker/__tests__/query.test.ts`
- Modify: `src/core/services/immersion-tracker/query.ts`
- Modify: `src/core/services/immersion-tracker/types.ts`

**Steps:**
1. Add red query tests for `word -> lines` and `kanji -> lines` mappings, including anime/video/session/timestamp/text context and per-line `occurrence_count`.
2. Run `bun test src/core/services/immersion-tracker/__tests__/query.test.ts` and confirm failure.
3. Implement the minimal query functions/types.
4. Re-run the targeted test and confirm green.

### Task 4: Expose the new query surface through the tracker service

**Files:**
- Modify: `src/core/services/immersion-tracker-service.ts`
- Modify: any narrow API/service consumer files only if needed

**Steps:**
1. Add the service methods needed to consume the new reverse-mapping queries.
2. Keep the change narrow; do not widen unrelated UI/API contracts unless a current consumer needs them.
3. Re-run the focused affected tests.

### Task 5: Verify with the maintained immersion lane

**Files:**
- Modify: `backlog/tasks/task-171 - Add-normalized-immersion-word-and-kanji-occurrence-tracking.md`

**Steps:**
1. Run the focused SQLite immersion tests first.
2. Escalate to broader verification only if touched files cross into API/runtime boundaries.
3. Record exact commands and results in the backlog task notes/final summary.
