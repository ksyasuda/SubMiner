---
id: TASK-351
title: Remove legacy global mpv plugin from setup
status: Done
assignee: []
created_date: '2026-05-12 19:57'
updated_date: '2026-05-12 20:03'
labels:
  - setup
  - mpv-plugin
  - launcher
  - windows
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add first-run setup support for detecting all legacy SubMiner mpv plugin auto-load entries and removing them via the OS trash after user confirmation, so regular mpv stops loading SubMiner while SubMiner-managed playback can use runtime plugin loading.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Setup detects all SubMiner mpv auto-load candidates under normal mpv scripts directories and Windows portable_config scripts directories.
- [x] #2 Setup displays detected legacy plugin paths and offers a Remove legacy mpv plugin action.
- [x] #3 Removal uses Electron shell.trashItem for detected script files/directories and never permanently deletes as fallback.
- [x] #4 script-opts/subminer.conf is not removed by the legacy plugin removal action.
- [x] #5 Partial trash failures report exact failed paths and keep legacy plugin warning visible.
- [x] #6 Successful removal refreshes setup status and reports that regular mpv will no longer load SubMiner while SubMiner-managed playback keeps working.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented setup detection for all legacy SubMiner mpv auto-load candidates in normal and portable mpv script directories, added a confirmed Remove legacy mpv plugin action that uses Electron shell.trashItem only, preserves script-opts/subminer.conf, reports exact partial failures, and refreshes setup status after successful removal. Added focused tests plus changelog/docs updates.
<!-- SECTION:FINAL_SUMMARY:END -->
