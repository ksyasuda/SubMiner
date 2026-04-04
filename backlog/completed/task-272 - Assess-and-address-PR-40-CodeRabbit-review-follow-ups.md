---
id: TASK-272
title: 'Assess and address PR #40 CodeRabbit review follow-ups'
status: Done
assignee: []
created_date: '2026-04-03 07:52'
updated_date: '2026-04-03 08:04'
labels:
  - coderabbit
  - review
  - launcher
milestone: 'PR #40'
dependencies: []
references:
  - 'https://github.com/ksyasuda/SubMiner/pull/40'
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Implement the valid CodeRabbit findings on PR #40 and keep the Windows mpv shortcut / first-run setup flow consistent end to end.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Windows binary resolution does not return install directories as executable candidates
- [ ] #2 Launch-mpv arg parsing preserves space-separated mpv option values and target separation
- [ ] #3 Windows mpv launch args keep the final input-ipc-server and script-opts socket path in sync when custom values are supplied
- [ ] #4 First-run setup navigation swallows stale or invalid custom-scheme actions without navigating away
- [ ] #5 Setup messaging and footer copy reflect configReady, plugin, and dictionary gates consistently
- [ ] #6 Regression tests cover the fixed behaviors
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Addressed CodeRabbit follow-ups for PR #40. Hardened launcher binary discovery on Windows and PATH resolution, fixed launch-mpv argument parsing for value-bearing flags, synced custom Windows mpv IPC socket values into script opts, and tightened first-run setup messaging/navigation to handle stale actions and blocker copy. Verified with `bun test src/main-entry-runtime.test.ts src/main/runtime/windows-mpv-launch.test.ts src/main/runtime/first-run-setup-window.test.ts launcher/mpv.test.ts`.
<!-- SECTION:FINAL_SUMMARY:END -->
