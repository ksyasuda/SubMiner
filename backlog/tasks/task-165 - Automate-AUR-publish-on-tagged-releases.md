---
id: TASK-165
title: Automate AUR publish on tagged releases
status: Done
assignee:
  - codex
created_date: '2026-03-14 15:55'
updated_date: '2026-03-14 18:40'
labels:
  - release
  - packaging
  - linux
dependencies:
  - TASK-161
references:
  - .github/workflows/release.yml
  - src/release-workflow.test.ts
  - docs/RELEASING.md
  - packaging/aur/subminer-bin/PKGBUILD
documentation:
  - docs/plans/2026-03-14-aur-release-sync-design.md
  - docs/plans/2026-03-14-aur-release-sync.md
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Extend the tagged release workflow so a successful GitHub release automatically updates the `subminer-bin` AUR package over SSH. Keep the PKGBUILD source-of-truth in this repo so release automation is reviewable and testable instead of depending on an external maintainer checkout.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Repo-tracked AUR packaging source exists for `subminer-bin` and matches the current release artifact layout.
- [x] #2 The release workflow clones `ssh://aur@aur.archlinux.org/subminer-bin.git` with a dedicated secret-backed SSH key only after release artifacts are ready.
- [x] #3 The workflow updates `pkgver`, regenerates `sha256sums` from the built release artifacts, regenerates `.SRCINFO`, and pushes only when packaging files changed.
- [x] #4 Regression coverage fails if the AUR publish job, secret contract, or update steps are removed from the release workflow.
- [x] #5 Release docs mention the required `AUR_SSH_PRIVATE_KEY` setup and the new tagged-release side effect.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Record the approved design and implementation plan for direct AUR publishing from the release workflow.
2. Add failing release workflow regression tests covering the new AUR publish job, SSH secret, and PKGBUILD/.SRCINFO regeneration steps.
3. Reintroduce repo-tracked `packaging/aur/subminer-bin` source files as the maintained AUR template.
4. Add a small helper script that updates `pkgver`, computes checksums from release artifacts, and regenerates `.SRCINFO` deterministically.
5. Extend `.github/workflows/release.yml` with an AUR publish job that clones the AUR repo over SSH, runs the helper, commits only when needed, and pushes to `aur`.
6. Update release docs for the new secret/setup requirements and tagged-release behavior.
7. Run targeted workflow tests plus the SubMiner verification lane needed for workflow/docs changes, then update this task with results.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added repo-tracked AUR packaging source under `packaging/aur/subminer-bin/` plus `scripts/update-aur-package.sh` to stamp `pkgver`, compute SHA-256 sums from release assets, and regenerate `.SRCINFO` with `makepkg --printsrcinfo`.

Extended `.github/workflows/release.yml` with a terminal `aur-publish` job that runs after `release`, validates `AUR_SSH_PRIVATE_KEY`, installs `makepkg`, configures SSH/known_hosts, clones `ssh://aur@aur.archlinux.org/subminer-bin.git`, downloads the just-published `SubMiner-<version>.AppImage`, `subminer`, and `subminer-assets.tar.gz` assets, updates packaging metadata, and pushes only when `PKGBUILD` or `.SRCINFO` changed.

Updated `src/release-workflow.test.ts` with regression assertions for the AUR publish contract and updated `docs/RELEASING.md` with the new secret/setup requirement.

Verification run:
- `bun test src/release-workflow.test.ts src/ci-workflow.test.ts`
- `bash -n scripts/update-aur-package.sh && bash -n packaging/aur/subminer-bin/PKGBUILD`
- `cd packaging/aur/subminer-bin && makepkg --printsrcinfo > .SRCINFO`
- updater smoke via temp package dir with fake assets and `v9.9.9`
- `bun run typecheck`
- `bun run test:fast`
- `bun run test:env`
- `git submodule update --init --recursive` (required because the worktree lacked release submodules)
- `bun run build`
- `bun run test:smoke:dist`

Docs update required: yes, completed in `docs/RELEASING.md`.
Changelog fragment required: no; internal release automation only.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Tagged releases now attempt a direct AUR sync for `subminer-bin` using a dedicated SSH private key stored in `AUR_SSH_PRIVATE_KEY`. The release workflow clones the AUR repo after GitHub Release publication, rewrites `PKGBUILD` and `.SRCINFO` from the published release assets, and skips empty pushes. Repo-owned packaging source and workflow regression coverage were added so the automation remains reviewable and testable.
<!-- SECTION:FINAL_SUMMARY:END -->
