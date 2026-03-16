---
id: TASK-161
title: Add Arch Linux PKGBUILD and .SRCINFO for SubMiner release artifacts
status: Done
assignee:
  - codex
created_date: '2026-03-11 07:50'
updated_date: '2026-03-16 05:13'
labels:
  - packaging
  - linux
  - docs
dependencies: []
references:
  - package.json
  - .github/workflows/release.yml
  - README.md
documentation:
  - docs-site/development.md
  - docs-site/installation.md
  - docs/RELEASING.md
priority: medium
ordinal: 27500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add repo-maintained Arch packaging metadata so Arch/AUR users can install the packaged SubMiner release artifacts without relying on npm. Cover the binary package flow that matches the current GitHub Releases distribution model and document the install path for Arch users.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 A dedicated Arch packaging directory exists with a PKGBUILD and matching .SRCINFO for the SubMiner binary release flow.
- [x] #2 The package installs the shipped Linux release artifact and wrapper in paths consistent with current runtime auto-detection expectations.
- [x] #3 Package metadata declares the required Arch runtime dependencies for the packaged workflow.
- [x] #4 The packaging files remain isolated from the main app/release/docs surfaces so they can be moved to a separate repository later.
- [x] #5 The packaging metadata is validated with an appropriate local Arch packaging check or an explicitly documented limitation if the tool is unavailable in this environment.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add isolated Arch packaging under packaging/arch/subminer-bin for the binary release flow that matches current GitHub Releases artifacts.
2. Install the Linux AppImage to /opt/SubMiner/SubMiner.AppImage, add PATH symlinks for SubMiner.AppImage and subminer, and ship optional plugin/theme/config assets under /usr/share/subminer.
3. Generate and commit matching .SRCINFO metadata in the same packaging directory.
4. Validate the PKGBUILD metadata locally with shell parsing plus makepkg --printsrcinfo and namcap if available.
5. Leave the new files isolated so they can be moved to a separate packaging repository later.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
User requested no commit and no broad repo integration; keep the Arch packaging files isolated in a dedicated directory for later extraction.

Created isolated Arch packaging under packaging/arch/subminer-bin with PKGBUILD and generated .SRCINFO only; no docs or release workflow changes.

Validation run: bash -n PKGBUILD, makepkg --printsrcinfo > .SRCINFO, namcap PKGBUILD.

PKGBUILD uses current v0.5.6 GitHub release assets and recorded SHA-256 values for SubMiner-0.5.6.AppImage, subminer, and subminer-assets.tar.gz.

Per user follow-up, moved packaging/arch/subminer-bin out of the repo to /home/sudacode/packages/maintaining/subminer-bin for separate maintenance. Repo docs/workflows were still left untouched.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added an isolated Arch Linux packaging directory at packaging/arch/subminer-bin containing a `subminer-bin` PKGBUILD and generated `.SRCINFO`. The package installs the current GitHub release AppImage to `/opt/SubMiner/SubMiner.AppImage`, adds PATH access for `SubMiner.AppImage` and `subminer`, and ships optional plugin/theme/config assets under `/usr/share/subminer`. Verified locally with `bash -n PKGBUILD`, `makepkg --printsrcinfo`, and `namcap PKGBUILD`.
<!-- SECTION:FINAL_SUMMARY:END -->
