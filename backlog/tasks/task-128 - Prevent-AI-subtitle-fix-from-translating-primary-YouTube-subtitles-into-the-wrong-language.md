---
id: TASK-128
title: >-
  Prevent AI subtitle fix from translating primary YouTube subtitles into the
  wrong language
status: Done
assignee: []
created_date: '2026-03-08 09:02'
updated_date: '2026-03-08 09:17'
labels:
  - bug
  - youtube-subgen
  - ai
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
AI subtitle cleanup can preserve cue structure while changing subtitle language, causing primary Japanese subtitle files to come back in English. Add guards so AI-fixed subtitles preserve expected language and fall back to raw Whisper output when language drifts.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Primary AI subtitle fix rejects output that drifts away from the expected source language.
- [x] #2 Rejected AI fixes fall back to the raw Whisper subtitle without corrupting published subtitle language.
- [x] #3 Regression tests cover a primary Japanese subtitle batch being translated into English by the AI fixer.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added a primary-language guard to AI subtitle fixing so Japanese source subtitles are rejected if the AI rewrites them into English while preserving SRT structure. The fixer now receives the expected source language from the YouTube orchestrator, and regression coverage verifies that language drift falls back to the raw Whisper subtitle path.
<!-- SECTION:FINAL_SUMMARY:END -->
