# Overlay Controller Support Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add Chrome Gamepad API controller support to the visible overlay as a supplement to keyboard-only mode, including controller selection/debug modals, config-backed logical bindings, and selected-token highlight cleanup.

**Architecture:** Keep controller support in the visible overlay renderer. Poll and normalize gamepad state in a dedicated runtime, route logical actions into the existing keyboard-only/Yomitan helpers, and persist preferred-controller config through the existing config pipeline and preload bridge.

**Tech Stack:** TypeScript, Bun tests, Electron preload IPC, renderer DOM modals, Chrome Gamepad API

---

### Task 1: Track work and lock the design

**Files:**
- Create: `backlog/tasks/task-159 - Add-overlay-controller-support-for-keyboard-only-mode.md`
- Create: `docs/plans/2026-03-11-overlay-controller-support-design.md`
- Create: `docs/plans/2026-03-11-overlay-controller-support.md`

**Step 1: Record the approved scope**

Capture controller-only-in-keyboard-mode behavior, the modal shortcuts, config scope, and the stale selection-highlight cleanup requirement.

**Step 2: Verify the written scope matches the approved design**

Run: `sed -n '1,220p' backlog/tasks/task-159\\ -\\ Add-overlay-controller-support-for-keyboard-only-mode.md && sed -n '1,240p' docs/plans/2026-03-11-overlay-controller-support-design.md`

Expected: task and design doc both mention controller selection/debug modals and highlight cleanup.

### Task 2: Add failing config tests and defaults

**Files:**
- Modify: `src/config/config.test.ts`
- Modify: `src/config/definitions/defaults-core.ts`
- Modify: `src/config/definitions/options-core.ts`
- Modify: `src/config/definitions/template-sections.ts`
- Modify: `src/types.ts`
- Modify: `config.example.jsonc`

**Step 1: Write the failing test**

Add coverage asserting a new `controller` config block resolves with the expected defaults and accepts logical-field overrides.

**Step 2: Run test to verify it fails**

Run: `bun test src/config/config.test.ts`

Expected: FAIL because `controller` config is not defined yet.

**Step 3: Write minimal implementation**

Add the controller config types/defaults/registry/template wiring and regenerate the example config if needed.

**Step 4: Run test to verify it passes**

Run: `bun test src/config/config.test.ts`

Expected: PASS

### Task 3: Add failing keyboard-selection cleanup tests

**Files:**
- Modify: `src/renderer/handlers/keyboard.test.ts`
- Modify: `src/renderer/handlers/keyboard.ts`
- Modify: `src/renderer/state.ts`

**Step 1: Write the failing tests**

Add tests for:

- turning keyboard-only mode off clears `.keyboard-selected`
- closing the popup clears stale selection highlight when keyboard-only mode is off

**Step 2: Run test to verify it fails**

Run: `bun test src/renderer/handlers/keyboard.test.ts`

Expected: FAIL because selection cleanup is incomplete today.

**Step 3: Write minimal implementation**

Centralize selection clearing in the keyboard-only sync helpers and popup-close flow.

**Step 4: Run test to verify it passes**

Run: `bun test src/renderer/handlers/keyboard.test.ts`

Expected: PASS

### Task 4: Add failing controller runtime tests

**Files:**
- Create: `src/renderer/handlers/gamepad-controller.test.ts`
- Create: `src/renderer/handlers/gamepad-controller.ts`
- Modify: `src/renderer/context.ts`
- Modify: `src/renderer/state.ts`
- Modify: `src/renderer/renderer.ts`

**Step 1: Write the failing tests**

Cover:

- first connected controller is selected by default
- preferred controller wins when connected
- controller actions are ignored unless keyboard-only mode is enabled, except keyboard-only toggle
- stick/button mappings invoke the expected logical helpers
- smooth scroll and repeat throttling behavior

**Step 2: Run test to verify it fails**

Run: `bun test src/renderer/handlers/gamepad-controller.test.ts`

Expected: FAIL because controller runtime does not exist.

**Step 3: Write minimal implementation**

Add a renderer-local polling runtime with deadzone handling, action edge detection, repeat timing, and helper callbacks into the keyboard/Yomitan flow.

**Step 4: Run test to verify it passes**

Run: `bun test src/renderer/handlers/gamepad-controller.test.ts`

Expected: PASS

### Task 5: Add failing controller modal tests

**Files:**
- Modify: `src/renderer/index.html`
- Modify: `src/renderer/style.css`
- Create: `src/renderer/modals/controller-select.ts`
- Create: `src/renderer/modals/controller-select.test.ts`
- Create: `src/renderer/modals/controller-debug.ts`
- Create: `src/renderer/modals/controller-debug.test.ts`
- Modify: `src/renderer/renderer.ts`
- Modify: `src/renderer/context.ts`
- Modify: `src/renderer/state.ts`

**Step 1: Write the failing tests**

Add tests for:

- `Alt+C` opens controller selection modal
- `Alt+Shift+C` opens controller debug modal
- selection modal renders connected controllers and persists the chosen device
- debug modal shows live axes/buttons state

**Step 2: Run test to verify it fails**

Run: `bun test src/renderer/modals/controller-select.test.ts src/renderer/modals/controller-debug.test.ts`

Expected: FAIL because modals and shortcuts do not exist.

**Step 3: Write minimal implementation**

Add modal DOM, renderer modules, modal state wiring, and controller runtime integration.

**Step 4: Run test to verify it passes**

Run: `bun test src/renderer/modals/controller-select.test.ts src/renderer/modals/controller-debug.test.ts`

Expected: PASS

### Task 6: Persist controller preference through preload/main wiring

**Files:**
- Modify: `src/preload.ts`
- Modify: `src/types.ts`
- Modify: `src/shared/ipc/contracts.ts`
- Modify: `src/core/services/ipc.ts`
- Modify: `src/main.ts`
- Modify: related main/runtime tests as needed

**Step 1: Write the failing test**

Add coverage for reading current controller config and saving preferred-controller changes from the renderer.

**Step 2: Run test to verify it fails**

Run: `bun test src/core/services/ipc.test.ts`

Expected: FAIL because no controller preference IPC exists yet.

**Step 3: Write minimal implementation**

Expose renderer-safe getters/setters for the controller config fields needed by the selection modal/runtime.

**Step 4: Run test to verify it passes**

Run: `bun test src/core/services/ipc.test.ts`

Expected: PASS

### Task 7: Update docs and config example

**Files:**
- Modify: `config.example.jsonc`
- Modify: `README.md`
- Modify: relevant docs under `docs-site/` for shortcuts/usage/troubleshooting if touched by current docs structure

**Step 1: Write the failing doc/config check if needed**

If config example generation is covered by tests, add/refresh the failing assertion first.

**Step 2: Implement the docs**

Document controller behavior, modal shortcuts, config block, and the keyboard-only-only activation rule.

**Step 3: Run doc/config verification**

Run: `bun run test:config`

Expected: PASS

### Task 8: Run the handoff gate and update the backlog task

**Files:**
- Modify: `backlog/tasks/task-159 - Add-overlay-controller-support-for-keyboard-only-mode.md`

**Step 1: Run targeted verification**

Run:

- `bun test src/config/config.test.ts`
- `bun test src/renderer/handlers/keyboard.test.ts`
- `bun test src/renderer/handlers/gamepad-controller.test.ts`
- `bun test src/renderer/modals/controller-select.test.ts`
- `bun test src/renderer/modals/controller-debug.test.ts`
- `bun test src/core/services/ipc.test.ts`

Expected: PASS

**Step 2: Run broader gate**

Run:

- `bun run typecheck`
- `bun run test:fast`
- `bun run test:env`
- `bun run build`

Expected: PASS, or document exact blockers/failures.

**Step 3: Update backlog notes**

Fill in implementation notes, verification commands, and final summary in `TASK-159`.
