---
id: TASK-106
title: Add first-run setup gate and auto-install flow
status: Done
assignee:
  - codex
created_date: '2026-03-07 06:10'
updated_date: '2026-03-16 05:13'
labels: []
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/main.ts
  - /home/sudacode/projects/japanese/SubMiner/src/shared/setup-state.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/runtime/first-run-setup-service.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/runtime/first-run-setup-window.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/launcher/commands/playback-command.ts
priority: high
ordinal: 13010
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Replace the current manual install flow with a first-run setup gate:

- bootstrap the default config dir/config file automatically
- detect legacy installs and mark them complete when config + Yomitan dictionaries are already present
- open a compact Catppuccin Macchiato setup popup for incomplete installs
- optionally install the mpv plugin into the default mpv location
- block launcher playback until setup completes, then resume the original playback flow
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 First app launch seeds the default config dir/config file without manual copy steps.
- [x] #2 Existing installs with config plus at least one Yomitan dictionary are auto-detected as already complete.
- [x] #3 Incomplete installs get a first-run setup popup with mpv plugin install, Yomitan settings, refresh, skip, and finish actions.
- [x] #4 Launcher playback waits for setup completion and does not start mpv while setup is incomplete.
- [x] #5 Plugin assets are packaged into the Electron bundle and regression tests cover the new flow.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added shared setup-state/config/mpv path helpers so Electron and launcher read the same onboarding state file.

Introduced a first-run setup service plus compact BrowserWindow popup using Catppuccin Macchiato styling. The popup supports optional mpv plugin install, opening Yomitan settings, status refresh, skip-plugin, and gated finish once at least one Yomitan dictionary is installed.

Electron startup now bootstraps a default config file, auto-detects legacy-complete installs, adds `--setup` CLI support, exposes a tray `Complete Setup` action while incomplete, and avoids reopening setup once completion is recorded.

Launcher playback now checks the shared setup-state file before starting mpv. If setup is incomplete, it launches the app with `--background --setup`, waits for completion, and only then proceeds.

Verification:

- `bun run typecheck`
- `bun run test:fast`
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
SubMiner now supports a download-and-launch install flow.

- First launch auto-creates config and opens setup only when needed.
- Existing users with working installs are silently migrated to completed setup.
- The setup popup handles optional mpv plugin install and Yomitan dictionary readiness.
- Launcher playback is gated on setup completion and resumes automatically afterward.
<!-- SECTION:FINAL_SUMMARY:END -->
