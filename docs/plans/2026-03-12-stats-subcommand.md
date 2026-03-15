# Stats Subcommand Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add `subminer stats` plus app-side `--stats` support so SubMiner can launch only the stats dashboard stack, reuse an existing app instance, and fail when immersion tracking is disabled.

**Architecture:** The launcher gets a new `stats` subcommand that forwards into a dedicated Electron CLI flag. The app handles that flag through a focused stats command service that validates immersion tracking, ensures tracker/server startup, opens the browser, and optionally reports success/failure through a response file so second-instance reuse preserves shell exit status.

**Tech Stack:** TypeScript, Bun test runner, Electron single-instance lifecycle, Commander-based launcher CLI.

---

### Task 1: Record launcher command coverage

**Files:**
- Modify: `launcher/main.test.ts`
- Modify: `launcher/commands/command-modules.test.ts`

**Step 1: Write the failing test**

Add coverage that `subminer stats` routes through a dedicated launcher command and forwards `--stats` into the app process.

**Step 2: Run test to verify it fails**

Run: `bun test launcher/main.test.ts launcher/commands/command-modules.test.ts`
Expected: FAIL because the `stats` command does not exist yet.

**Step 3: Write minimal implementation**

Add the launcher parser, command module, and dispatch wiring needed for the tests to pass.

**Step 4: Run test to verify it passes**

Run: `bun test launcher/main.test.ts launcher/commands/command-modules.test.ts`
Expected: PASS

**Step 5: Commit**

```bash
git add launcher/main.test.ts launcher/commands/command-modules.test.ts launcher/config/cli-parser-builder.ts launcher/main.ts launcher/commands/stats-command.ts
git commit -m "feat: add launcher stats command"
```

### Task 2: Add failing app CLI parsing/runtime tests

**Files:**
- Modify: `src/cli/args.test.ts`
- Modify: `src/main/runtime/cli-command-runtime-handler.test.ts`
- Modify: `src/main/runtime/cli-command-context.test.ts`

**Step 1: Write the failing test**

Add tests for:
- parsing `--stats`
- starting the app for `--stats`
- stats command behavior when immersion tracking is disabled
- stats command behavior when startup succeeds

**Step 2: Run test to verify it fails**

Run: `bun test src/cli/args.test.ts src/main/runtime/cli-command-runtime-handler.test.ts src/main/runtime/cli-command-context.test.ts`
Expected: FAIL because the new flag/service is not implemented.

**Step 3: Write minimal implementation**

Extend CLI args and add the smallest stats command runtime surface required by the tests.

**Step 4: Run test to verify it passes**

Run: `bun test src/cli/args.test.ts src/main/runtime/cli-command-runtime-handler.test.ts src/main/runtime/cli-command-context.test.ts`
Expected: PASS

**Step 5: Commit**

```bash
git add src/cli/args.test.ts src/main/runtime/cli-command-runtime-handler.test.ts src/main/runtime/cli-command-context.test.ts src/cli/args.ts src/main/cli-runtime.ts src/main/runtime/cli-command-context.ts src/main/runtime/cli-command-context-deps.ts
git commit -m "feat: add app stats cli command"
```

### Task 3: Add stats-only startup plumbing

**Files:**
- Modify: `src/main.ts`
- Modify: `src/core/services/startup.ts`
- Modify: `src/core/services/cli-command.ts`
- Modify: `src/main/runtime/immersion-startup.ts`
- Test: existing runtime tests above plus any new focused stats tests

**Step 1: Write the failing test**

Add focused tests around the stats startup service and response-path reporting before touching production logic.

**Step 2: Run test to verify it fails**

Run: `bun test src/core/services/cli-command.test.ts`
Expected: FAIL because stats startup/reporting is missing.

**Step 3: Write minimal implementation**

Implement:
- stats-only startup gating
- tracker/server/browser startup orchestration
- response-file success/failure reporting for reused primary-instance handling

**Step 4: Run test to verify it passes**

Run: `bun test src/core/services/cli-command.test.ts`
Expected: PASS

**Step 5: Commit**

```bash
git add src/main.ts src/core/services/startup.ts src/core/services/cli-command.ts src/main/runtime/immersion-startup.ts src/core/services/cli-command.test.ts
git commit -m "feat: add stats-only startup flow"
```

### Task 4: Update docs/help

**Files:**
- Modify: `src/cli/help.ts`
- Modify: `docs-site/immersion-tracking.md`
- Modify: `docs-site/mining-workflow.md`

**Step 1: Write the failing test**

Add/update any help-text assertions first.

**Step 2: Run test to verify it fails**

Run: `bun test src/cli/help.test.ts`
Expected: FAIL until help text includes stats usage.

**Step 3: Write minimal implementation**

Document `subminer stats` and clarify that explicit invocation forces the local dashboard server to start.

**Step 4: Run test to verify it passes**

Run: `bun test src/cli/help.test.ts`
Expected: PASS

**Step 5: Commit**

```bash
git add src/cli/help.ts src/cli/help.test.ts docs-site/immersion-tracking.md docs-site/mining-workflow.md
git commit -m "docs: document stats subcommand"
```

### Task 5: Verify integrated behavior

**Files:**
- No new files required

**Step 1: Run targeted unit/integration lanes**

Run:
- `bun run typecheck`
- `bun run test:launcher`
- `bun test src/cli/args.test.ts src/core/services/cli-command.test.ts src/main/runtime/cli-command-runtime-handler.test.ts src/main/runtime/cli-command-context.test.ts src/cli/help.test.ts`

Expected: PASS

**Step 2: Run broader maintained lane if the targeted slice is clean**

Run: `bun run test:fast`
Expected: PASS or surface unrelated pre-existing failures.

**Step 3: Commit verification fixes if needed**

```bash
git add -A
git commit -m "test: stabilize stats command coverage"
```
