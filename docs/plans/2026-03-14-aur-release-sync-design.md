# AUR Release Sync Design

**Date:** 2026-03-14
**Status:** Approved

## Goal

Publish `subminer-bin` to the AUR automatically when a tagged SubMiner release completes successfully.

## Chosen Approach

Keep the `subminer-bin` packaging template in this repo under `packaging/aur/subminer-bin/`. The tag-driven GitHub Actions release workflow will reuse the release artifacts it already built, clone `ssh://aur@aur.archlinux.org/subminer-bin.git` over SSH, rewrite `PKGBUILD` and `.SRCINFO` from the current tag and artifact checksums, then push only if the resulting packaging files changed.

## Why This Approach

- Keeps AUR metadata reviewable in the main repo instead of hiding it in workflow YAML.
- Avoids depending on `/home/sudacode/packages/maintaining/subminer-bin`, which CI cannot access.
- Lets `src/release-workflow.test.ts` enforce the workflow contract with simple text assertions.
- Keeps the release side effect sequenced after the existing GitHub Release publish step, so AUR only points at artifacts that already exist publicly.

## Scope

- Add repo-owned `PKGBUILD` template and generated `.SRCINFO` under `packaging/aur/subminer-bin/`.
- Add a helper script to stamp `pkgver`, rewrite `sha256sums`, and regenerate `.SRCINFO`.
- Add a release workflow job that:
  - uses `AUR_SSH_PRIVATE_KEY`
  - populates `known_hosts` for `aur.archlinux.org`
  - clones the AUR repo
  - copies in the repo-owned packaging source
  - updates packaging metadata from release artifacts
  - commits only when files changed
  - pushes to the `aur` remote
- Update release docs for secret setup and behavior.

## Non-Goals

- Managing AUR state from a separate maintainer mirror.
- Publishing non-binary Arch packages.
- Triggering AUR updates outside tagged releases.

## Risks And Guards

- Secret missing or invalid: fail fast in the AUR publish job with a clear error.
- Artifact drift: compute checksums from the artifacts downloaded in the same workflow run.
- Spurious empty commits: guard with `git diff --quiet`.
- Workflow regression: cover the job shape in `src/release-workflow.test.ts`.
