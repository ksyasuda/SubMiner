---
id: TASK-162
title: Normalize packaged Linux paths to canonical SubMiner directories
status: In Progress
assignee:
  - codex
created_date: '2026-03-11 08:28'
updated_date: '2026-03-11 08:29'
labels:
  - linux
  - packaging
  - docs
dependencies: []
references:
  - launcher/mpv.ts
  - launcher/picker.ts
  - plugin/subminer/binary.lua
  - plugin/subminer.conf
  - docs-site/installation.md
  - docs-site/launcher-script.md
  - README.md
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Align packaged Linux path conventions so system-installed assets use canonical `SubMiner` directories and match runtime auto-detection. Cover AppImage binary discovery, rofi theme discovery/docs, and related path references while preserving lowercase names only for the launcher wrapper, rofi theme filename, and mpv Lua plugin/conf.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Launcher/runtime path discovery prefers canonical packaged Linux locations that use `SubMiner` casing for shared data and config directories.
- [ ] #2 Tests cover the expected packaged Linux discovery paths for the AppImage and rofi theme search behavior.
- [ ] #3 User-facing docs reference the canonical packaged Linux locations consistently.
- [ ] #4 Lowercase names remain only where intentionally required for the launcher wrapper, rofi theme filename, and mpv Lua plugin/conf.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing launcher tests for canonical packaged Linux discovery paths: /usr/lib/subminer/SubMiner.AppImage via PATH symlink flow and /usr/share/SubMiner/themes/subminer.rasi for rofi theme lookup.
2. Update launcher runtime path discovery to prefer canonical packaged Linux shared-data locations using SubMiner casing.
3. Update plugin auto-detection comments and binary search defaults so packaged Linux paths stay consistent with launcher/runtime expectations.
4. Update user-facing docs to reference canonical SubMiner-cased config/share paths while keeping lowercase names only for the launcher wrapper, rofi theme filename, and mpv Lua plugin/conf.
5. Run targeted launcher tests plus docs checks.
<!-- SECTION:PLAN:END -->
