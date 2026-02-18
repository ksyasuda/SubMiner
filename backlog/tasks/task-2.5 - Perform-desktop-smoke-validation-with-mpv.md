---
id: TASK-2.5
title: Perform desktop smoke validation with mpv
status: Done
assignee: []
created_date: '2026-02-10 18:56'
updated_date: '2026-02-18 04:11'
labels: []
dependencies:
  - TASK-2.2
  - TASK-2.3
  - TASK-2.4
references:
  - investigation.md
  - docs/refactor-main-checklist.md
parent_task_id: TASK-2
ordinal: 13000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Execute manual desktop smoke checks in an MPV-enabled environment to validate overlay rendering and key user workflows not fully covered by automated tests.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Overlay rendering and visibility toggling are verified in a real desktop session.
- [x] #2 Card mining flow is verified end-to-end.
- [x] #3 Field-grouping interaction is verified end-to-end.
- [x] #4 Results and any follow-up defects are documented in task notes.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Smoke run executed on 2026-02-10 with real Electron launch (outside sandbox) after unsetting `ELECTRON_RUN_AS_NODE=1` in command context.

Commands executed:

- `electron . --help`
- `electron . --start`
- `electron . --toggle-visible-overlay`
- `electron . --toggle-invisible-overlay`
- `electron . --mine-sentence`
- `electron . --trigger-field-grouping`
- `electron . --open-runtime-options`
- `electron . --stop`

Observed runtime evidence from app logs:

- CLI help output rendered with expected flags.
- App started and connected to MPV after reconnect attempts.
- Mining flow executed and produced `Created sentence card: ...`, plus media upload logs.
- Tracker/runtime loop started (`hyprland` tracker connected) and app stopped cleanly.

Follow-up/constraints:

- Overlay _visual rendering_ and visibility correctness are not directly observable from terminal logs alone and still require direct desktop visual confirmation.
- Field-grouping trigger command was sent, but explicit end-state confirmation in UI still needs manual verification.
<!-- SECTION:NOTES:END -->
