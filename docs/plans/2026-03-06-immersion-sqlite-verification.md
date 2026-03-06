# Immersion SQLite Verification Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Make the SQLite-backed immersion tracking persistence tests visible in the repo's verification surface and reproducible through at least one documented automated command.

**Architecture:** Keep the existing Bun fast lane intact for routine local verification, but add an explicit SQLite verification lane that runs the database-backed immersion tests under a runtime with `node:sqlite` support. Surface unsupported-runtime behavior clearly in the source tests and contributor docs so skipped or omitted coverage is no longer mistaken for a fully green persistence lane.

**Tech Stack:** TypeScript, Bun scripts in `package.json`, Node's built-in `node:test` and `node:sqlite`, GitHub Actions workflows, Markdown docs in `README.md`.

---

### Task 1: Audit and expose the SQLite-backed immersion test surface

**Files:**

- Modify: `src/core/services/immersion-tracker-service.test.ts`
- Modify: `src/core/services/immersion-tracker/storage-session.test.ts`
- Reference: `src/main/runtime/registry.test.ts`

**Step 1: Write the failing test**

Refactor the SQLite-gated immersion tests so missing `node:sqlite` support is reported with an explicit skip reason instead of a silent top-level `test.skip` alias.

**Step 2: Run test to verify it fails**

Run: `bun test src/core/services/immersion-tracker-service.test.ts src/core/services/immersion-tracker/storage-session.test.ts`
Expected: the current output shows generic skips or hides the storage-session suite from normal scripted verification, which is too opaque for contributors.

**Step 3: Write minimal implementation**

Mirror the `src/main/runtime/registry.test.ts` pattern: add a helper that either loads `DatabaseSync` or skips with a message like `requires node:sqlite support in this runtime`, then wrap each SQLite-backed test through that helper.

**Step 4: Run test to verify it passes**

Run: `bun test src/core/services/immersion-tracker-service.test.ts src/core/services/immersion-tracker/storage-session.test.ts`
Expected: PASS, with explicit skip messages in unsupported runtimes.

### Task 2: Add a reproducible SQLite verification command

**Files:**

- Modify: `package.json`
- Reference: `src/core/services/immersion-tracker-service.test.ts`
- Reference: `src/core/services/immersion-tracker/storage-session.test.ts`

**Step 1: Write the failing test**

Add a dedicated script contract for the SQLite-backed immersion verification lane so both persistence-heavy suites are intentionally grouped and runnable together.

**Step 2: Run test to verify it fails**

Run: `bun run test:immersion:sqlite`
Expected: FAIL because no such reproducible lane exists yet.

**Step 3: Write minimal implementation**

Update `package.json` with explicit scripts for the SQLite lane. Prefer a command shape that actually executes the built JS tests under Node with `node:sqlite` support, for example:

- `test:immersion:sqlite:dist`: `node --test dist/core/services/immersion-tracker-service.test.js dist/core/services/immersion-tracker/storage-session.test.js`
- `test:immersion:sqlite`: `bun run build && bun run test:immersion:sqlite:dist`

If build cost or runtime behavior requires a small adjustment, keep the core contract the same: one documented command must run both SQLite-backed immersion suites end-to-end.

**Step 4: Run test to verify it passes**

Run: `bun run test:immersion:sqlite`
Expected: PASS in a Node runtime with `node:sqlite`, executing both persistence suites without Bun-only skips.

### Task 3: Wire the SQLite lane into automated verification

**Files:**

- Modify: `.github/workflows/ci.yml`
- Modify: `.github/workflows/release.yml`
- Reference: `package.json`

**Step 1: Write the failing test**

Add the new SQLite immersion lane to the repo's automated verification so contributors and CI can rely on a real persistence check rather than the Bun fast lane alone.

**Step 2: Run test to verify it fails**

Run: `bun run test:immersion:sqlite`
Expected: local command may pass, but CI/release workflows still omit the lane entirely.

**Step 3: Write minimal implementation**

Update both workflows to provision a Node version with `node:sqlite` support before the SQLite lane runs, then execute `bun run test:immersion:sqlite` in the quality gate after the bundle build produces `dist/**` test files.

**Step 4: Run test to verify it passes**

Run: `bun run test:immersion:sqlite`
Expected: PASS locally, and workflow definitions clearly show the SQLite lane as part of automated verification.

### Task 4: Document contributor-facing prerequisites and commands

**Files:**

- Modify: `README.md`
- Reference: `package.json`
- Reference: `.github/workflows/ci.yml`

**Step 1: Write the failing test**

Extend the verification docs so contributors can discover the SQLite lane, know why the Bun source lane may skip those cases, and understand which command reproduces the persistence coverage.

**Step 2: Run test to verify it fails**

Run: `grep -n "test:immersion:sqlite" README.md`
Expected: FAIL because the dedicated immersion SQLite lane is undocumented.

**Step 3: Write minimal implementation**

Update `README.md` to document:

- the Bun fast/default lane versus the SQLite persistence lane
- the `node:sqlite` prerequisite for the reproducible command
- that the dedicated lane covers session persistence/finalization behavior beyond seam tests

**Step 4: Run test to verify it passes**

Run: `grep -n "test:immersion:sqlite" README.md && grep -n "node:sqlite" README.md`
Expected: PASS, with clear contributor guidance.

### Task 5: Verify persistence coverage end-to-end

**Files:**

- Test: `src/core/services/immersion-tracker-service.test.ts`
- Test: `src/core/services/immersion-tracker/storage-session.test.ts`
- Reference: `README.md`
- Reference: `package.json`

**Step 1: Write the failing test**

Prove the final lane exercises real DB-backed persistence/finalization paths, not just the seam tests.

**Step 2: Run test to verify it fails**

Run: `bun run test:immersion:sqlite`
Expected: before implementation, the command does not exist or does not cover both SQLite-backed suites.

**Step 3: Write minimal implementation**

Keep the dedicated lane pointed at both existing SQLite-backed test files so it covers representative finalization and persistence behavior such as:

- `destroy finalizes active session and persists final telemetry`
- `start/finalize session updates ended_at and status`
- `executeQueuedWrite inserts event and telemetry rows`

**Step 4: Run test to verify it passes**

Run: `bun run test:immersion:sqlite`
Expected: PASS, with those DB-backed persistence/finalization cases executing successfully under Node.
