---
id: TASK-118
title: Add Windows release build and SignPath signing
status: Done
assignee:
  - codex
created_date: '2026-03-08 15:17'
updated_date: '2026-03-16 05:13'
labels:
  - release
  - windows
  - signing
dependencies: []
references:
  - .github/workflows/release.yml
  - build/installer.nsh
  - build/signpath-windows-artifact-config.xml
  - package.json
priority: high
ordinal: 54500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Extend the tag-driven release workflow so Windows artifacts are built on GitHub-hosted runners and submitted to SignPath for free open-source Authenticode signing, while preserving the existing macOS notarization path.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Release workflow builds Windows installer and ZIP artifacts on `windows-latest`
- [x] #2 Workflow submits unsigned Windows artifacts to SignPath and uploads the signed outputs for release publication
- [x] #3 Repository includes a checked-in SignPath artifact-configuration source of truth for the Windows release files
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Inspect the existing release workflow and current Windows packaging configuration.
2. Add a Windows release job that builds unsigned artifacts, uploads them as a workflow artifact, and submits them to SignPath.
3. Update the release aggregation job to publish signed Windows assets and mention Windows install steps in the generated release notes.
4. Check in the Windows SignPath artifact configuration XML used to define what gets signed.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
The repository already had Windows packaging configuration (`build:win`, NSIS include script, Windows helper asset packaging), but the release workflow still built Linux and macOS only.

Added a `build-windows` job to `.github/workflows/release.yml` that runs on `windows-latest`, validates required SignPath secrets, builds unsigned Windows artifacts, uploads them with `actions/upload-artifact@v4`, and then calls the official `signpath/github-action-submit-signing-request@v2` action to retrieve signed outputs.

Checked in `build/signpath-windows-artifact-config.xml` as the source-of-truth artifact configuration for SignPath. It signs the top-level NSIS installer EXE and deep-signs `.exe` and `.dll` files inside the portable ZIP artifact.

Updated the release aggregation job to download the signed Windows artifacts and added a Windows install section to the generated GitHub release body.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Windows release publishing is now wired into the tag-driven workflow. `.github/workflows/release.yml` builds Windows artifacts on `windows-latest`, submits them to SignPath using the official GitHub action, and publishes the signed `.exe` and `.zip` outputs alongside the Linux and macOS artifacts. The workflow now requests the additional `actions: read` permission required by the SignPath GitHub integration, and the generated release notes now include Windows installation steps.

The checked-in `build/signpath-windows-artifact-config.xml` file defines the SignPath artifact structure expected by the workflow artifact ZIP: sign the top-level `SubMiner-*.exe` installer and deep-sign `.exe` and `.dll` files inside `SubMiner-*.zip`.

Verification: workflow/static changes were checked with `git diff --check` on the touched files. Actual signing requires configured SignPath secrets and a matching artifact configuration in your SignPath project.
<!-- SECTION:FINAL_SUMMARY:END -->
