---
id: TASK-126
title: >-
  Improve secondary subtitle readability with hover-only background and stronger
  text separation
status: Done
assignee: []
created_date: '2026-03-08 07:35'
updated_date: '2026-03-08 07:40'
labels:
  - overlay
  - subtitles
  - ui
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Adjust overlay secondary subtitle styling so translation text stays readable on bright video backgrounds. Keep the dark background hidden by default in hover mode and show it only while hovered. Increase secondary subtitle weight to 600 and strengthen edge separation without changing primary subtitle styling.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Secondary subtitles render with stronger edge separation than today.
- [x] #2 Secondary subtitle font weight defaults to 600.
- [x] #3 When secondary subtitle mode is hover, the secondary background appears only while hovered.
- [x] #4 Primary subtitle styling behavior remains unchanged.
- [x] #5 Renderer tests cover the new secondary hover background behavior and default secondary style values.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Adjusted secondary subtitle defaults to use stronger shadowing, 600 font weight, and a translucent dark background. Routed secondary background/backdrop styling through CSS custom properties so hover mode can keep the background hidden until the secondary subtitle is actually hovered. Added renderer and config tests covering default values and hover-only background behavior.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Improved secondary subtitle readability by strengthening default text separation, increasing the default secondary weight to 600, and making the configured dark background appear only while hovered in secondary hover mode. Added config and renderer coverage for the new defaults and hover-aware style routing.
<!-- SECTION:FINAL_SUMMARY:END -->
