# Merged Character Dictionary Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Replace per-anime character dictionary imports with one merged Yomitan dictionary driven by MRU usage retention.

**Architecture:** Persist normalized per-media character dictionary snapshots locally, maintain MRU retained media ids in auto-sync state, and rebuild a single merged Yomitan zip only when the retained set changes. Keep external AniList fetches only for media without a local snapshot; normal revisits stay local.

**Tech Stack:** TypeScript, Bun test, Node fs/path, existing Yomitan zip generation helpers.

---

### Task 1: Lock in merged auto-sync behavior

**Files:**
- Modify: `src/main/runtime/character-dictionary-auto-sync.test.ts`
- Test: `src/main/runtime/character-dictionary-auto-sync.test.ts`

**Step 1: Write the failing test**

Add tests for:
- single merged dictionary title/import replacing per-media imports
- MRU reorder causing rebuild only when order changes
- unchanged revisit skipping rebuild/import
- capped retained set evicting least-recently-used media

**Step 2: Run test to verify it fails**

Run: `bun test src/main/runtime/character-dictionary-auto-sync.test.ts`
Expected: FAIL on old per-media import assumptions / missing merged behavior

**Step 3: Write minimal implementation**

Update auto-sync runtime to track retained media ids and merged revision/hash, call merged zip builder, and replace one imported Yomitan dictionary.

**Step 4: Run test to verify it passes**

Run: `bun test src/main/runtime/character-dictionary-auto-sync.test.ts`
Expected: PASS

### Task 2: Add snapshot + merged-zip runtime support

**Files:**
- Modify: `src/main/character-dictionary-runtime.ts`
- Modify: `src/main/character-dictionary-runtime.test.ts`
- Test: `src/main/character-dictionary-runtime.test.ts`

**Step 1: Write the failing test**

Add tests for:
- saving/loading normalized per-media snapshots without per-media zip cache
- building merged zip from retained media snapshots with stable dictionary title
- preserving images/terms from multiple media in merged output

**Step 2: Run test to verify it fails**

Run: `bun test src/main/character-dictionary-runtime.test.ts`
Expected: FAIL because snapshot/merged APIs do not exist yet

**Step 3: Write minimal implementation**

Refactor dictionary runtime to expose snapshot generation/loading and merged zip building from stored metadata/images.

**Step 4: Run test to verify it passes**

Run: `bun test src/main/character-dictionary-runtime.test.ts`
Expected: PASS

### Task 3: Wire app/runtime config and docs

**Files:**
- Modify: `src/main.ts`
- Modify: `src/config/definitions/options-integrations.ts`
- Modify: `README.md`

**Step 1: Write the failing test**

Add or update tests if needed for new dependency wiring / docs-adjacent config description expectations.

**Step 2: Run test to verify it fails**

Run: `bun test src/main/runtime/character-dictionary-auto-sync.test.ts src/main/character-dictionary-runtime.test.ts`
Expected: FAIL until wiring matches merged flow

**Step 3: Write minimal implementation**

Swap app wiring to new snapshot + merged build API, update config/docs text from TTL semantics to usage-based merged retention.

**Step 4: Run test to verify it passes**

Run: `bun test src/main/runtime/character-dictionary-auto-sync.test.ts src/main/character-dictionary-runtime.test.ts && bun run tsc --noEmit`
Expected: PASS
