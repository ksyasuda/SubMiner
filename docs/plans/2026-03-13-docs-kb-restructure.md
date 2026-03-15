# Internal Knowledge Base Restructure Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Turn `AGENTS.md` into a compact table of contents, establish `docs/` as the internal system of record, keep `docs-site/` user-facing, and enforce the split with tests/CI.

**Architecture:** Create a small internal knowledge-base hierarchy under `docs/`, migrate canonical contributor/agent guidance into it, and wire a repo-level verification test that checks required docs, metadata, and boundary rules. Keep public docs in `docs-site/` and update them to reference internal docs rather than acting as the source of truth themselves.

**Tech Stack:** Markdown docs, Bun test runner, existing GitHub Actions CI

---

### Task 1: Add the internal KB entrypoints

**Files:**
- Create: `docs/README.md`
- Create: `docs/architecture/README.md`
- Create: `docs/architecture/domains.md`
- Create: `docs/architecture/layering.md`
- Create: `docs/knowledge-base/README.md`
- Create: `docs/knowledge-base/core-beliefs.md`
- Create: `docs/knowledge-base/catalog.md`
- Create: `docs/knowledge-base/quality.md`
- Create: `docs/workflow/README.md`
- Create: `docs/workflow/planning.md`
- Create: `docs/workflow/verification.md`

**Step 1: Write the KB home page**

Add `docs/README.md` with navigation to architecture, workflow, knowledge-base maintenance, release docs, and active plans.

**Step 2: Add architecture pages**

Create the architecture index plus focused `domains.md` and `layering.md` pages that summarize runtime ownership and dependency boundaries from the existing architecture doc.

**Step 3: Add knowledge-base pages**

Create the knowledge-base index, core beliefs, catalog, and quality pages with explicit metadata fields and short maintenance guidance.

**Step 4: Add workflow pages**

Create workflow index, planning guide, and verification guide with the current maintained Bun commands and lane-selection guidance.

**Step 5: Review cross-links**

Read the new docs and confirm every page links back to at least one parent/index page.

### Task 2: Rewrite `AGENTS.md` as a compact map

**Files:**
- Modify: `AGENTS.md`

**Step 1: Replace encyclopedia-style content with a compact map**

Keep only the minimum operational guidance needed in injected context:
- quick start
- internal source-of-truth pointers
- build/test gate
- generated/sensitive file notes
- release pointer
- backlog note

**Step 2: Add direct links to the new KB entrypoints**

Point `AGENTS.md` at:
- `docs/README.md`
- `docs/architecture/README.md`
- `docs/workflow/README.md`
- `docs/workflow/verification.md`
- `docs/knowledge-base/README.md`
- `docs/RELEASING.md`

**Step 3: Keep the file intentionally short**

Target roughly 100 lines and avoid moving deep details back into `AGENTS.md`.

### Task 3: Re-boundary `docs-site/`

**Files:**
- Modify: `docs-site/development.md`
- Modify: `docs-site/architecture.md`
- Modify: `docs-site/README.md`

**Step 1: Update contributor-facing docs**

Keep build/run/testing instructions, but stop presenting `docs-site/*` pages as canonical internal architecture/workflow references.

**Step 2: Add explicit internal-doc pointers**

Link readers to `docs/README.md` and the new internal architecture/workflow pages for deep contributor guidance.

**Step 3: Preserve public-doc tone**

Ensure the `docs-site/` pages remain user/contributor-facing and do not become the internal KB themselves.

### Task 4: Add mechanical enforcement

**Files:**
- Create: `scripts/docs-knowledge-base.test.ts`
- Modify: `package.json`
- Modify: `.github/workflows/ci.yml`

**Step 1: Write a repo-level docs KB test**

Assert:
- required docs exist
- metadata fields exist on internal docs
- `AGENTS.md` links to internal KB entrypoints
- `AGENTS.md` stays under the line cap
- `docs-site/development.md` and `docs-site/architecture.md` point to `docs/`

**Step 2: Wire the test into package scripts**

Add a script for the KB test and include it in an existing maintained verification lane.

**Step 3: Ensure CI exercises the check**

Make sure the CI path that runs the maintained test lane catches KB regressions.

### Task 5: Verify and hand off

**Files:**
- Modify: any files above if verification reveals drift

**Step 1: Run targeted docs verification**

Run:
- `bun run docs:test`
- `bun run test:fast`
- `bun run docs:build`

**Step 2: Fix drift found by tests**

If any assertions fail, update the docs or test expectations so the enforced model matches the intended structure.

**Step 3: Summarize outcome**

Report the new internal KB entrypoints, the `AGENTS.md` table-of-contents rewrite, enforcement coverage, verification results, and any skipped items.
