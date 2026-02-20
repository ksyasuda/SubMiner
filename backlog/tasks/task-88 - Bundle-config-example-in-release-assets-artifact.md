---
id: TASK-88
title: Bundle config.example.jsonc in release assets artifact
status: Done
assignee: []
created_date: '2026-02-20 09:27'
updated_date: '2026-02-20 09:27'
labels: []
dependencies: []
priority: medium
ordinal: 70000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Release users should get a version-matched `config.example.jsonc` without cloning the repo. Include the file in `subminer-assets-*.tar.gz` and align install snippets/docs.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Release workflow includes `config.example.jsonc` in `subminer-assets-*.tar.gz`.
- [x] #2 Release notes optional-assets section includes config copy step.
- [x] #3 Installation docs show config copy from extracted release assets bundle.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Updated `.github/workflows/release.yml` optional assets packaging to include `config.example.jsonc`, and updated release body instructions accordingly. Updated `docs/installation.md` and `docs/mpv-plugin.md` release-bundle install snippets to copy `/tmp/config.example.jsonc` into `~/.config/SubMiner/config.jsonc`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
`config.example.jsonc` is now bundled in the release `subminer-assets-*.tar.gz` artifact, with release/install docs aligned so users can copy it directly after extracting the bundle.
<!-- SECTION:FINAL_SUMMARY:END -->
