# Character Name Gating Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Disable subtitle character-name lookup/highlighting when the AniList character dictionary feature is disabled, while keeping tokenization and all other annotations working.

**Architecture:** Gate `getNameMatchEnabled` at the runtime-deps boundary used by subtitle tokenization. Keep the tokenizer pipeline intact and only suppress character-name metadata requests when `anilist.characterDictionary.enabled` is false, regardless of `subtitleStyle.nameMatchEnabled`.

**Tech Stack:** TypeScript, Bun test runner, Electron main/runtime wiring.

---

### Task 1: Add runtime gating coverage

**Files:**
- Modify: `src/main/runtime/subtitle-tokenization-main-deps.test.ts`

**Step 1: Write the failing test**

Add a test proving `getNameMatchEnabled()` resolves to `false` when `getCharacterDictionaryEnabled()` is `false` even if `getNameMatchEnabled()` is `true`.

**Step 2: Run test to verify it fails**

Run: `bun test src/main/runtime/subtitle-tokenization-main-deps.test.ts`
Expected: FAIL because the deps builder does not yet combine the two flags.

### Task 2: Implement minimal runtime gate

**Files:**
- Modify: `src/main/runtime/subtitle-tokenization-main-deps.ts`
- Modify: `src/main.ts`

**Step 3: Write minimal implementation**

Add `getCharacterDictionaryEnabled` to the main handler deps and make the built `getNameMatchEnabled` return true only when both the subtitle setting and the character dictionary setting are enabled.

**Step 4: Run tests to verify green**

Run: `bun test src/main/runtime/subtitle-tokenization-main-deps.test.ts`
Expected: PASS.

### Task 3: Verify no regressions in related tokenization seams

**Files:**
- Modify: none unless failures reveal drift

**Step 5: Run focused verification**

Run: `bun test src/core/services/subtitle-processing-controller.test.ts src/main/runtime/subtitle-tokenization-main-deps.test.ts`
Expected: PASS.
