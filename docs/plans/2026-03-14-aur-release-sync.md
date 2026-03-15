# AUR Release Sync Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Automatically update the `subminer-bin` AUR package when a tagged GitHub release completes.

**Architecture:** Keep `subminer-bin` packaging metadata in-repo and use a small deterministic helper script to stamp version and checksums from the release artifacts. Extend the existing tag-driven release workflow with a final SSH-based AUR publish job that runs after GitHub Release publication.

**Tech Stack:** GitHub Actions, Bash, makepkg, Bun workflow tests, PKGBUILD/.SRCINFO

---

### Task 1: Record work and lock the approved design

**Files:**
- Create: `backlog/tasks/task-165 - Automate-AUR-publish-on-tagged-releases.md`
- Create: `docs/plans/2026-03-14-aur-release-sync-design.md`
- Create: `docs/plans/2026-03-14-aur-release-sync.md`

**Step 1: Record the approved scope**

Capture direct AUR publishing, the in-repo packaging source of truth, the `AUR_SSH_PRIVATE_KEY` secret contract, and the release-artifact checksum flow.

**Step 2: Verify the written scope**

Run: `sed -n '1,220p' backlog/tasks/task-165\\ -\\ Automate-AUR-publish-on-tagged-releases.md && sed -n '1,220p' docs/plans/2026-03-14-aur-release-sync-design.md`

Expected: both files mention direct AUR SSH push and repo-tracked packaging source.

### Task 2: Add failing release workflow tests

**Files:**
- Modify: `src/release-workflow.test.ts`

**Step 1: Write the failing tests**

Add assertions that require:
- an `aur-publish` release job
- `AUR_SSH_PRIVATE_KEY`
- `ssh://aur@aur.archlinux.org/subminer-bin.git`
- `makepkg --printsrcinfo`
- a guard that avoids empty commit/push runs

**Step 2: Run test to verify it fails**

Run: `bun test src/release-workflow.test.ts`

Expected: FAIL because the current release workflow has no AUR publish job.

### Task 3: Add repo-tracked AUR packaging source and updater

**Files:**
- Create: `packaging/aur/subminer-bin/PKGBUILD`
- Create: `packaging/aur/subminer-bin/.SRCINFO`
- Create: `scripts/update-aur-package.sh`

**Step 1: Write the failing updater expectation if needed**

If helper behavior is complex enough, add a focused test or keep coverage in the release workflow assertions plus shell validation.

**Step 2: Write minimal implementation**

Add a checked-in PKGBUILD template and a helper that stamps `pkgver`, computes sha256 sums from release artifacts, and regenerates `.SRCINFO`.

**Step 3: Run focused verification**

Run:
- `bash -n scripts/update-aur-package.sh`
- `bash -n packaging/aur/subminer-bin/PKGBUILD`

Expected: PASS

### Task 4: Extend the release workflow

**Files:**
- Modify: `.github/workflows/release.yml`

**Step 1: Implement the AUR publish job**

Add a final release job that:
- depends on `release`
- installs `makepkg`
- configures SSH from `AUR_SSH_PRIVATE_KEY`
- clones the AUR repo
- copies in packaging source
- runs the updater
- commits only if files changed
- pushes to the AUR remote

**Step 2: Run workflow regression tests**

Run: `bun test src/release-workflow.test.ts`

Expected: PASS

### Task 5: Update release docs

**Files:**
- Modify: `docs/RELEASING.md`

**Step 1: Document the new release side effect**

Add the secret/setup requirement and note that tagged releases now attempt an AUR sync.

**Step 2: Run docs/workflow verification**

Run:
- `bun test src/release-workflow.test.ts`
- `bun test src/ci-workflow.test.ts`

Expected: PASS

### Task 6: Run handoff verification and update backlog notes

**Files:**
- Modify: `backlog/tasks/task-165 - Automate-AUR-publish-on-tagged-releases.md`

**Step 1: Run targeted verification**

Run:
- `bun test src/release-workflow.test.ts`
- `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane docs`

Expected: PASS, or exact blocker captured.

**Step 2: Run broader required gate for touched areas**

Run:
- `bun run typecheck`

Expected: PASS

**Step 3: Update task notes**

Record implementation notes, verification commands, docs decision, and final summary.
