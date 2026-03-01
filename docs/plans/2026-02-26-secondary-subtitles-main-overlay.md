# Secondary Subtitles Main Overlay Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Ensure secondary subtitles render in the unified main overlay window and remove stale secondary-window/layer paths.

**Architecture:** Keep secondary subtitle DOM in the shared renderer tree, rely on mode classes (`secondary-sub-hidden|visible|hover`) for visibility, and remove obsolete legacy overlay-layer assumptions. Preserve modal behavior and existing subtitle rendering flow.

**Tech Stack:** TypeScript, Electron renderer CSS/DOM, Bun test runner.

---

### Task 1: Add Regression Tests For Main Overlay Secondary Rendering

**Files:**

- Modify: `src/renderer/subtitle-render.test.ts`
- Modify: `src/renderer/error-recovery.test.ts`

**Step 1: Write failing tests**

- Assert stylesheet no longer hides secondary subtitles in `layer-visible`.
- Assert renderer platform resolution ignores legacy `secondary` overlay layer.

**Step 2: Run tests to verify failures**

Run: `bun test src/renderer/subtitle-render.test.ts src/renderer/error-recovery.test.ts`
Expected: FAIL on secondary subtitle hide rule + legacy secondary layer handling.

### Task 2: Remove Secondary-Window CSS/Routing Assumptions

**Files:**

- Modify: `src/renderer/style.css`
- Modify: `src/renderer/utils/platform.ts`
- Modify: `src/renderer/error-recovery.ts`
- Modify: `src/types.ts`

**Step 1: Implement minimal changes**

- Remove legacy forced hide on `#secondarySubContainer`.
- Remove obsolete layer-specific secondary-subtitle CSS blocks.
- Drop legacy `secondary` overlay-layer parsing path from renderer platform resolver.
- Narrow related overlay layer type unions.

**Step 2: Run targeted tests**

Run: `bun test src/renderer/subtitle-render.test.ts src/renderer/error-recovery.test.ts`
Expected: PASS.

### Task 3: Validate Wider Related Surface

**Files:**

- No additional code changes required.

**Step 1: Run broader related tests**

Run: `bun test src/renderer/subtitle-render.test.ts src/renderer/error-recovery.test.ts src/main/runtime/overlay-window-runtime-handlers.test.ts src/main/runtime/overlay-window-factory.test.ts src/core/services/overlay-manager.test.ts`
Expected: Renderer tests pass; report any unrelated pre-existing failures.
