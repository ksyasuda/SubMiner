# Unsigned Windows Release Builds Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Publish unsigned Windows release artifacts in GitHub Actions while adding an explicit local `build:win:unsigned` script.

**Architecture:** Keep Windows packaging on `electron-builder`, but stop the release workflow from routing artifacts through SignPath. The Windows release job will build unsigned artifacts and upload them directly under the existing `windows` artifact name so the downstream release job stays stable. Local developer behavior remains unchanged except for a new explicit unsigned build script.

**Tech Stack:** GitHub Actions, Bun, Electron Builder, Node test runner

---

### Task 1: Track the workflow contract change

**Files:**
- Create: `backlog/tasks/task-138 - Publish-unsigned-Windows-release-artifacts-and-add-local-unsigned-build-script.md`
- Create: `changes/unsigned-windows-release-builds.md`

**Step 1: Write the backlog task + changelog fragment**

Document the scope: unsigned Windows release CI, new local unsigned script, no SignPath dependency.

**Step 2: Review file formatting**

Run: `sed -n '1,220p' backlog/tasks/task-138\ -\ Publish-unsigned-Windows-release-artifacts-and-add-local-unsigned-build-script.md && sed -n '1,80p' changes/unsigned-windows-release-builds.md`
Expected: task metadata matches existing backlog files; changelog fragment matches `changes/README.md` format.

### Task 2: Write failing workflow regression tests

**Files:**
- Modify: `src/release-workflow.test.ts`

**Step 1: Write the failing test**

Replace SignPath-specific workflow assertions with assertions for:
- unsigned Windows artifacts built via `bun run build:win:unsigned`
- direct `windows` artifact upload from `release/*.exe` and `release/*.zip`
- no SignPath action references
- package scripts include `build:win:unsigned`

**Step 2: Run test to verify it fails**

Run: `bun test src/release-workflow.test.ts`
Expected: FAIL because the current workflow still validates SignPath secrets and submits signing requests.

### Task 3: Patch package scripts and release workflow

**Files:**
- Modify: `package.json`
- Modify: `.github/workflows/release.yml`

**Step 1: Write minimal implementation**

- add `build:win:unsigned` that clears Windows signing env and disables auto discovery before invoking `electron-builder --win nsis zip --publish never`
- change the Windows release job to remove SignPath secret validation/submission
- build Windows artifacts with `bun run build:win:unsigned`
- upload `release/*.exe` and `release/*.zip` directly as `windows`

**Step 2: Run tests to verify they pass**

Run: `bun test src/release-workflow.test.ts`
Expected: PASS

### Task 4: Run focused verification

**Files:**
- Modify: none

**Step 1: Run focused checks**

Run: `bun test src/release-workflow.test.ts && bun run typecheck`
Expected: all green

**Step 2: Spot-check diff**

Run: `git --no-pager diff -- .github/workflows/release.yml package.json src/release-workflow.test.ts changes/unsigned-windows-release-builds.md backlog/tasks/task-138\ -\ Publish-unsigned-Windows-release-artifacts-and-add-local-unsigned-build-script.md docs/plans/2026-03-09-unsigned-windows-release-builds.md`
Expected: only scoped unsigned-Windows workflow/script/docs changes
