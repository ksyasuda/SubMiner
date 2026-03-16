---
id: TASK-162
title: Normalize packaged Linux paths to canonical SubMiner directories
status: Done
assignee:
  - codex
created_date: '2026-03-11 08:28'
updated_date: '2026-03-16 06:25'
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
- [x] #1 Launcher/runtime path discovery prefers canonical packaged Linux locations that use `SubMiner` casing for shared data and config directories.
- [x] #2 Tests cover the expected packaged Linux discovery paths for the AppImage and rofi theme search behavior.
- [x] #3 User-facing docs reference the canonical packaged Linux locations consistently.
- [x] #4 Lowercase names remain only where intentionally required for the launcher wrapper, rofi theme filename, and mpv Lua plugin/conf.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing launcher tests for canonical packaged Linux discovery paths: /usr/lib/subminer/SubMiner.AppImage via PATH symlink flow and /usr/share/SubMiner/themes/subminer.rasi for rofi theme lookup.
2. Update launcher runtime path discovery to prefer canonical packaged Linux shared-data locations using SubMiner casing.
3. Update plugin auto-detection comments and binary search defaults so packaged Linux paths stay consistent with launcher/runtime expectations.
4. Update user-facing docs to reference canonical SubMiner-cased config/share paths while keeping lowercase names only for the launcher wrapper, rofi theme filename, and mpv Lua plugin/conf.
5. Run targeted launcher tests plus docs checks.

Remaining work (2026-03-15):
- binary.lua: add lowercase fallback candidates /usr/bin/subminer and /usr/local/bin/subminer after existing title-case entries
- launcher tests: add findAppBinary Linux candidates and findRofiTheme /usr/share + /usr/local/share tests
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-03-15: Adding launcher tests for Linux packaged path discovery (findAppBinary + findRofiTheme). Implementing in mpv.test.ts and new picker.test.ts following node:test / assert/strict patterns from mpv.test.ts.

2026-03-15: AC#2 complete. Added findAppBinary tests (3) to launcher/mpv.test.ts and findRofiTheme tests (4) to new launcher/picker.test.ts. All 76 launcher tests pass. Added picker.test.ts to test:launcher:src script.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
## Completed changes\n\n### `plugin/subminer/binary.lua`\nAdded lowercase fallback candidates after existing title-case entries in the non-Windows `find_binary()` search list:\n- `/usr/local/bin/subminer` (after `/usr/local/bin/SubMiner`)\n- `/usr/bin/subminer` (after `/usr/bin/SubMiner`)\n\n### `plugin/subminer.conf`\nUpdated the comment documenting the Linux binary search list to include the two new lowercase candidates.\n\n### `launcher/mpv.test.ts`\nAdded 3 new tests for `findAppBinary` Linux candidates:\n- Resolves `~/.local/bin/SubMiner.AppImage` when it exists\n- Resolves `/opt/SubMiner/SubMiner.AppImage` when `~/.local/bin` candidate absent\n- Finds `subminer` on PATH when AppImage candidates absent\n\n### `launcher/picker.test.ts` (new file)\nAdded 4 tests for `findRofiTheme` Linux packaged paths:\n- Resolves `/usr/local/share/SubMiner/themes/subminer.rasi`\n- Resolves `/usr/share/SubMiner/themes/subminer.rasi` when `/usr/local/share` absent\n- Resolves `$XDG_DATA_HOME/SubMiner/themes/subminer.rasi` when set\n- Resolves `~/.local/share/SubMiner/themes/subminer.rasi` when `XDG_DATA_HOME` unset\n\n### `package.json`\nAdded `launcher/picker.test.ts` to `test:launcher:src` file list.\n\n## Verification\n- `launcher-plugin` lane: passed (76 launcher tests, 524 fast tests — all green)\n\n## Policy checks\n- Docs update required? No — docs already reflected canonical paths.\n- Changelog fragment required? Yes — user-visible fix to plugin binary auto-detection. Fragment should be added under `changes/`.
<!-- SECTION:FINAL_SUMMARY:END -->
