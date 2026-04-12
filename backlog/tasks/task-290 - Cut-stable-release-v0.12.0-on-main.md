---
id: TASK-290
title: Cut stable release v0.12.0 on main
status: Done
assignee:
  - codex
created_date: '2026-04-12 04:47'
updated_date: '2026-04-12 04:51'
labels: []
dependencies: []
documentation:
  - docs/RELEASING.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Prepare the main branch for the stable SubMiner v0.12.0 release by applying the release-version updates, formatting changes required by the branch state, and rerunning the full release verification gate.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Main branch version and stable release metadata are updated for v0.12.0.
- [x] #2 Required formatting changes for the release candidate tree are applied and verified.
- [x] #3 The documented release verification gate passes locally and any remaining push or tag prerequisites are documented.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Audit main-branch release state: package version, release artifacts, current CI status, and current formatting debt.
2. Apply required formatting fixes to the files reported by `bun run format:check:src` and verify the formatting lane passes.
3. Update the package version to 0.12.0 and generate stable release metadata (`CHANGELOG.md`, `release/release-notes.md`, `docs-site/changelog.md`) using the documented release workflow.
4. Run the full local release gate on main (`changelog:lint`, `changelog:check --version 0.12.0`, `verify:config-example`, `typecheck`, `test:fast`, `test:env`, `build`, `docs:test`, `docs:build`, plus dist smoke) and document any remaining tag/push prerequisites.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Applied Prettier to all 39 files reported by `bun run format:check:src` on main and verified the formatting lane now passes.

Reapplied the stable changelog build entrypoint fix on main: added `writeStableReleaseArtifacts`, covered it with a focused regression test, and updated `package.json` so `changelog:build` forwards `--version` and `--date` through a single `build-release` command.

Verified the formatted mainline release tree with `bun run changelog:lint`, `bun run changelog:check --version 0.12.0`, `bun run verify:config-example`, `bun run typecheck`, `bun run test:fast`, `bun run test:env`, `bun run build`, `bun run docs:test`, `bun run docs:build`, and `bun run test:smoke:dist`; all passed.

Remote main CI also completed successfully for `Windows update (#49)` after the local release-prep pass. Remaining operational steps are commit/tag/push only.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Prepared `main` for the stable `v0.12.0` cut. Formatted the previously failing source files so `bun run format:check:src` is now clean, bumped `package.json` from `0.12.0-beta.3` to `0.12.0`, and generated the stable release artifacts with the explicit local cut date `2026-04-11`, which consumed the pending changelog fragments into `CHANGELOG.md`, `docs-site/changelog.md`, and `release/release-notes.md`.

Also reintroduced the release-script fix on main: the old `changelog:build` package script still used `build && docs`, which can drop `--version/--date` on the first step. Added a focused regression test in `scripts/build-changelog.test.ts`, implemented `writeStableReleaseArtifacts` in `scripts/build-changelog.ts`, and switched `package.json` to `build-release` so release flags propagate correctly. Verification on the final tree passed for formatting, changelog lint/check, config example verification, typecheck, fast tests, env tests, build, docs tests/build, dist smoke, and remote main CI. The branch is release-ready pending commit, tag, and push.
<!-- SECTION:FINAL_SUMMARY:END -->
