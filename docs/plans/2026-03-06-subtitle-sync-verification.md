# Subtitle Sync Verification Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Replace the no-op `test:subtitle` lane with real automated subtitle-sync verification that reuses the maintained subsync tests and documents the real contributor workflow.

**Architecture:** Repoint the subtitle verification command at the existing source-level subsync tests instead of inventing a second hidden suite. Add one focused ffsubsync failure-path test so the subtitle lane explicitly covers both engines plus a non-happy path, then update contributor docs to describe the dedicated subtitle lane and how it relates to `test:core`.

**Tech Stack:** TypeScript, Bun test, Node test/assert, npm package scripts, Markdown docs.

---

### Task 1: Lock subtitle lane to real subsync tests

**Files:**

- Modify: `package.json`

**Step 1: Write the failing test**

Define the intended command shape first: `test:subtitle:src` should run `src/core/services/subsync.test.ts` and `src/subsync/utils.test.ts`, `test:subtitle` should invoke that real source lane, and no placeholder echo should remain.

**Step 2: Run test to verify it fails**

Run: `bun run test:subtitle`
Expected: It performs a build and prints `Subtitle tests are currently not configured`, proving the lane is still a no-op.

**Step 3: Write minimal implementation**

Update `package.json` so:

- `test:subtitle:src` runs `bun test src/core/services/subsync.test.ts src/subsync/utils.test.ts`
- `test:subtitle` runs the new source lane directly
- `test:subtitle:dist` is removed if it is no longer the real verification path

**Step 4: Run test to verify it passes**

Run: `bun run test:subtitle`
Expected: PASS with Bun executing the real subtitle-sync test files.

### Task 2: Add explicit ffsubsync non-happy-path coverage

**Files:**

- Modify: `src/core/services/subsync.test.ts`
- Test: `src/core/services/subsync.test.ts`

**Step 1: Write the failing test**

Add a test that runs `runSubsyncManual({ engine: 'ffsubsync' })` with a stub ffsubsync executable that exits non-zero and writes stderr, then assert:

- `result.ok === false`
- `result.message` starts with `ffsubsync synchronization failed`
- the failure message includes command details surfaced to the user

**Step 2: Run test to verify it fails**

Run: `bun test src/core/services/subsync.test.ts`
Expected: FAIL because ffsubsync failure propagation is not asserted yet.

**Step 3: Write minimal implementation**

Keep production code unchanged unless the new test exposes a real bug. If needed, tighten failure assertions or message propagation in `src/core/services/subsync.ts` without changing successful behavior.

**Step 4: Run test to verify it passes**

Run: `bun test src/core/services/subsync.test.ts`
Expected: PASS with both alass and ffsubsync paths covered, including a non-happy path.

### Task 3: Make contributor docs match the real verification path

**Files:**

- Modify: `README.md`
- Modify: `package.json`

**Step 1: Write the failing test**

Use the repository state as the failure signal: README currently advertises subtitle sync as a feature but does not tell contributors that `bun run test:subtitle` is the real verification lane.

**Step 2: Run test to verify it fails**

Run: `bun run test:subtitle && bun test src/subsync/utils.test.ts`
Expected: Tests pass, but docs still do not explain the lane; this is the remaining acceptance-criteria gap.

**Step 3: Write minimal implementation**

Update `README.md` with a short contributor-facing verification note that:

- points to `bun run test:subtitle` for subtitle-sync coverage
- states that the lane reuses the maintained subsync tests already included in broader core coverage
- avoids implying there is a separate hidden subtitle test harness beyond those tests

**Step 4: Run test to verify it passes**

Run: `bun run test:subtitle`
Expected: PASS, with docs and scripts now aligned around the same subtitle verification strategy.

### Task 4: Verify matrix integration stays clean

**Files:**

- Modify: `package.json` (only if Task 1/3 exposed cleanup needs)

**Step 1: Write the failing test**

Treat duplication as the failure condition: confirm the dedicated subtitle lane reuses the same maintained files already present in `test:core:src` rather than creating a second divergent suite.

**Step 2: Run test to verify it fails**

Run: `bun run test:subtitle && bun run test:core:src`
Expected: If file lists diverge unexpectedly, this review step exposes it before handoff.

**Step 3: Write minimal implementation**

If needed, do the smallest script cleanup necessary so subtitle coverage remains explicit without hiding or duplicating existing core coverage.

**Step 4: Run test to verify it passes**

Run: `bun run test:subtitle && bun run test:core:src`
Expected: PASS, confirming the dedicated lane and the broader core suite agree on subtitle coverage.
