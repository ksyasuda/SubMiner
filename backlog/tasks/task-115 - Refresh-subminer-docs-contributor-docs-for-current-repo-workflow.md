---
id: TASK-115
title: Refresh subminer-docs contributor docs for current repo workflow
status: Done
assignee:
  - codex
created_date: '2026-03-08 00:40'
updated_date: '2026-03-08 00:42'
labels:
  - docs
dependencies: []
references:
  - ../subminer-docs/development.md
  - ../subminer-docs/README.md
  - Makefile
  - package.json
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Update the sibling `subminer-docs` repo so contributor/development docs match the current SubMiner repo workflow after the docs split and recent tooling changes, including removing stale in-repo docs build steps and documenting the scoped formatting command.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Contributor docs in `subminer-docs` no longer reference stale in-repo docs build commands for the app repo
- [x] #2 Contributor docs mention the current scoped formatting workflow (`make pretty` / `format:src`) where relevant
- [x] #3 Removed stale or no-longer-needed instructions that no longer match the current repo layout
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Inspect `subminer-docs` for contributor/development instructions that drifted after the docs repo split and recent tooling changes.
2. Update contributor docs to remove stale app-repo docs commands and document the current scoped formatting workflow.
3. Verify the modified docs page and build the docs site from the sibling docs repo when local dependencies are available.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Detected concrete doc drift in `subminer-docs/development.md`: stale in-repo docs build commands and no mention of the scoped `make pretty` formatter.

Updated `../subminer-docs/development.md` to remove stale app-repo docs build steps from the local gate, document `make pretty` / `format:check:src`, and point docs-site work to the sibling docs repo explicitly.

Installed docs repo dependencies locally with `bun install` and verified the docs site with `bun run docs:build` in `../subminer-docs`.

Did not change `../subminer-docs/README.md`; it was already accurate for the docs repo itself.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Refreshed the contributor/development docs in the sibling `subminer-docs` repo to match the current SubMiner workflow. In `development.md`, removed the stale app-repo `bun run docs:build` step from the local CI-equivalent gate, added an explicit note to run docs builds from `../subminer-docs` when docs change, documented the scoped formatting workflow (`make pretty` and `bun run format:check:src`), and replaced the old in-repo `make docs*` instructions with the correct sibling-repo `bun run docs:*` commands. Also updated the Makefile reference to include `make pretty` and removed the obsolete `make docs-dev` entry.

Verification: installed docs repo dependencies with `bun install` in `../subminer-docs` and ran `bun run docs:build` successfully. Left `README.md` unchanged because it was already accurate for the standalone docs repo.
<!-- SECTION:FINAL_SUMMARY:END -->
