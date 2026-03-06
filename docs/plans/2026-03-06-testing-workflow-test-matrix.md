# Testing Workflow Test Matrix Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Make the standard test commands reflect the maintained test surface so newly added tests are discovered automatically or intentionally documented outside the default lane.

**Architecture:** Replace the current hand-maintained file allowlists in `package.json` with directory-based Bun test lanes that map to maintained test surfaces. Keep the default developer lane fast, move slower or environment-specific checks into explicit commands, and document the resulting matrix in `README.md` so contributors know exactly which command to run.

**Tech Stack:** TypeScript, Bun test, npm-style package scripts in `package.json`, Markdown docs in `README.md`.

---

### Task 1: Lock in the desired script matrix with failing tests/audit checks

**Files:**

- Modify: `package.json`
- Test: `package.json`
- Reference: `src/main-entry-runtime.test.ts`
- Reference: `src/anki-integration/anki-connect-proxy.test.ts`
- Reference: `src/main/runtime/registry.test.ts`

**Step 1: Write the failing test**

Add a new script structure in `package.json` expectations by editing the script map so these lanes exist conceptually:

- `test:fast` for default fast verification
- `test:full` for the maintained source test surface
- `test:env` for environment-specific checks

The fast lane should stay selective and intentional. The full lane should use directory-based discovery rather than file-by-file allowlists, with representative coverage from:

- `src/main-entry-runtime.test.ts`
- `src/anki-integration/**/*.test.ts`
- `src/main/**/*.test.ts`
- `launcher/**/*.test.ts`

**Step 2: Run test to verify it fails**

Run: `bun run test:full`
Expected: FAIL because `test:full` does not exist yet, and previously omitted maintained tests are still outside the standard matrix.

**Step 3: Write minimal implementation**

Update `package.json` scripts so:

- `test` points at `test:fast`
- `test:fast` runs the fast default lane only
- `test:full` runs directory-based maintained suites instead of file allowlists
- `test:env` runs environment-specific verification (for example launcher/plugin and sqlite-gated suites)
- subsystem scripts use stable path globs or directory arguments so new tests are discovered automatically

Prefer commands like these, adjusted only as needed for Bun behavior in this repo:

- `bun test src/config/**/*.test.ts`
- `bun test src/{cli,core,renderer,subtitle,subsync,main,anki-integration}/*.test.ts ...` only if Bun cannot take the broader directory directly
- `bun test launcher/**/*.test.ts`

Do not keep large hand-maintained file enumerations for maintained unit/integration lanes.

**Step 4: Run test to verify it passes**

Run: `bun run test:full`
Expected: PASS, including automated execution of representative tests that were previously omitted from the standard matrix.

### Task 2: Separate environment-specific verification from the maintained default/full lanes

**Files:**

- Modify: `package.json`
- Test: `src/main/runtime/registry.test.ts`
- Test: `launcher/smoke.e2e.test.ts`
- Test: `src/core/services/immersion-tracker-service.test.ts`

**Step 1: Write the failing test**

Refine the package scripts so environment-specific checks are explicitly grouped outside the default fast lane. Treat these as the primary environment-specific examples unless repo behavior proves a better split during execution:

- launcher smoke/plugin checks that rely on local process or Lua execution
- sqlite-dependent checks that may skip when `node:sqlite` is unavailable

**Step 2: Run test to verify it fails**

Run: `bun run test:env`
Expected: FAIL because the environment-specific lane is not defined yet.

**Step 3: Write minimal implementation**

Add explicit environment-specific scripts in `package.json`, such as:

- a launcher/plugin lane that runs `launcher/smoke.e2e.test.ts` plus `lua scripts/test-plugin-start-gate.lua`
- a sqlite lane for tests that require `node:sqlite` support or otherwise need environment notes
- an aggregate `test:env` command that runs all environment-specific lanes

Keep these lanes documented and reproducible rather than silently excluded.

**Step 4: Run test to verify it passes**

Run: `bun run test:env`
Expected: PASS in supported environments, or clear documented skip behavior where the tests themselves intentionally gate on missing runtime support.

### Task 3: Document contributor-facing test commands and matrix

**Files:**

- Modify: `README.md`
- Reference: `package.json`

**Step 1: Write the failing test**

Add a contributor-focused testing section requirement in `README.md` expectations:

- fast verification command
- full verification command
- environment-specific verification command
- plain-language explanation of which suites each lane covers and why

**Step 2: Run test to verify it fails**

Run: `grep -n "Testing" README.md`
Expected: no contributor testing matrix section exists yet.

**Step 3: Write minimal implementation**

Update `README.md` with a concise `Testing` section that documents:

- `bun run test` / `bun run test:fast` for fast local verification
- `bun run test:full` for the maintained source test surface
- `bun run test:env` for environment-specific verification
- any important notes about sqlite-gated tests and launcher/plugin checks

Keep the matrix concrete and reproducible.

**Step 4: Run test to verify it passes**

Run: `grep -n "Testing" README.md && grep -n "test:full" README.md && grep -n "test:env" README.md`
Expected: PASS with the new contributor-facing matrix present.

### Task 4: Verify representative omitted suites now belong to automated lanes

**Files:**

- Test: `src/main-entry-runtime.test.ts`
- Test: `src/anki-integration/anki-connect-proxy.test.ts`
- Test: `src/main/runtime/registry.test.ts`
- Reference: `package.json`
- Reference: `README.md`

**Step 1: Write the failing test**

Use targeted command checks to prove these previously omitted surfaces are now in the matrix:

- entry/runtime: `src/main-entry-runtime.test.ts`
- Anki integration: `src/anki-integration/anki-connect-proxy.test.ts`
- main runtime: `src/main/runtime/registry.test.ts`

**Step 2: Run test to verify it fails**

Run: `bun run test:full src/main-entry-runtime.test.ts`
Expected: either unsupported invocation or evidence that the current matrix still does not include these surfaces automatically.

**Step 3: Write minimal implementation**

Adjust the final script paths/globs until the full matrix includes those representative surfaces without file-by-file script maintenance.

**Step 4: Run test to verify it passes**

Run: `bun test src/main-entry-runtime.test.ts src/anki-integration/anki-connect-proxy.test.ts src/main/runtime/registry.test.ts && bun run test:fast && bun run test:full`
Expected: PASS, with at least one representative test from each required surface executing through the documented automated lanes.
