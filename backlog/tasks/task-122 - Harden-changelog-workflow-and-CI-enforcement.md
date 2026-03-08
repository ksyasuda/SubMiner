---
id: TASK-122
title: Harden changelog workflow and CI enforcement
status: Done
assignee:
  - Codex
created_date: '2026-03-08 06:13'
updated_date: '2026-03-08 06:28'
labels:
  - release
  - changelog
  - ci
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/scripts/build-changelog.ts
  - /Users/sudacode/projects/japanese/SubMiner/scripts/build-changelog.test.ts
  - /Users/sudacode/projects/japanese/SubMiner/.github/workflows/ci.yml
  - /Users/sudacode/projects/japanese/SubMiner/.github/workflows/release.yml
  - /Users/sudacode/projects/japanese/SubMiner/docs/RELEASING.md
  - /Users/sudacode/projects/japanese/SubMiner/changes/README.md
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Improve the release changelog workflow so changelog fragments are reliable, release output is more readable, and pull requests get early feedback when changelog metadata is missing or malformed.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `scripts/build-changelog.ts` ignores non-fragment files in `changes/` and validates fragment structure before generating changelog output.
- [x] #2 Generated `CHANGELOG.md` and `release/release-notes.md` group public changes into readable sections instead of a flat bullet list.
- [x] #3 CI enforces changelog validation on pull requests and provides an explicit opt-out path for changes that should not produce release notes.
- [x] #4 Contributor docs explain the fragment format and the PR/release workflow for changelog generation.
- [x] #5 Automated tests cover fragment parsing/building behavior and workflow enforcement expectations.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing tests for changelog fragment discovery, structured fragment parsing/rendering, release-note output, and CI workflow expectations.
2. Update scripts/build-changelog.ts to ignore non-fragment files, parse fragment metadata, group generated output by change type, add lint/PR-check commands, and simplify output paths to repo-local artifacts.
3. Update CI and PR workflow files to run changelog validation on pull requests with an explicit skip path, and keep release workflow using committed changelog output.
4. Refresh changes/README.md, docs/RELEASING.md, and any PR template text so contributors know how to write fragments and when opt-out is allowed.
5. Run targeted tests and changelog commands, then record results and finalize the task.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented structured changelog fragments with required `type` and `area` metadata; `changes/README.md` is now ignored by the generator and verified by regression tests.

Added `changelog:lint` and `changelog:pr-check`, plus PR CI enforcement with `skip-changelog` opt-out. PR check now reads git name-status output so deleted fragment files do not satisfy the requirement.

Changed generated changelog/release notes output to grouped sections (`Added`, `Changed`, `Fixed`, etc.) and simplified release notes to highlights + install/assets pointers.

Kept changelog output repo-local. This aligns with existing repo direction where docs updates happen in the sibling docs repo explicitly rather than implicit local writes from app-repo generators.

Verification: `bun test scripts/build-changelog.test.ts src/ci-workflow.test.ts src/release-workflow.test.ts` passed; `bun run typecheck` passed; `bun run changelog:lint` passed. `bun run test:fast` still fails in unrelated existing `src/core/services/subsync.test.ts` cases (`runSubsyncManual keeps internal alass source file alive until sync finishes`, `runSubsyncManual resolves string sid values from mpv stream properties`).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Hardened the changelog workflow end-to-end. `scripts/build-changelog.ts` now ignores helper files like `changes/README.md`, requires structured fragment metadata (`type` + `area`), groups generated release sections by change type, and emits shorter release notes focused on highlights plus install/assets pointers. Added explicit `changelog:lint` and `changelog:pr-check` commands, with PR validation based on git name-status so deleted fragment files do not satisfy the fragment requirement.

Updated contributor-facing workflow docs in `changes/README.md`, `docs/RELEASING.md`, and a new PR template so authors know to add a fragment or apply the `skip-changelog` label. CI now runs fragment linting on every run and enforces fragment presence on pull requests. Added regression coverage in `scripts/build-changelog.test.ts` and a new `src/ci-workflow.test.ts` to lock the workflow contract.

Verification completed: `bun test scripts/build-changelog.test.ts src/ci-workflow.test.ts src/release-workflow.test.ts`, `bun run typecheck`, and `bun run changelog:lint` all passed. A broader `bun run test:fast` run still fails in unrelated existing `src/core/services/subsync.test.ts` cases outside the changelog/workflow scope.
<!-- SECTION:FINAL_SUMMARY:END -->
