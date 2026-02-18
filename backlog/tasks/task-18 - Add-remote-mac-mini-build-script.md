---
id: TASK-18
title: Add remote mac mini build script
status: Done
assignee:
  - codex
created_date: '2026-02-11 16:48'
updated_date: '2026-02-18 04:11'
labels:
  - build
  - macos
  - devex
dependencies: []
references:
  - scripts/build-external.sh
priority: medium
ordinal: 54000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add a script that offloads macOS package builds to a remote Mac mini over SSH, then syncs release artifacts back to the local workspace.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Script supports remote host and path configuration.
- [x] #2 Script supports signed and unsigned macOS build modes.
- [x] #3 Script syncs project sources to remote and copies release artifacts back locally.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added `scripts/build-external.sh` with:
- Defaults for host alias (`mac-mini`) and remote path (`~/build/SubMiner`)
- Argument flags: `--host`, `--remote-path`, `--signed`, `--unsigned`, `--skip-sync`, `--sync-only`
- Rsync-based upload excluding heavyweight build artifacts and local dependencies
- Remote build execution (`pnpm run build:mac` or `pnpm run build:mac:unsigned`)
- Artifact sync back into local `release/`
<!-- SECTION:NOTES:END -->
