---
id: TASK-70
title: Polish YouTube subtitle generation pipeline
status: To Do
assignee: []
created_date: '2026-02-26 07:37'
labels: []
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
$Current YouTube subtitle generation in launcher/youtube.ts is functional but has implicit behavior, low observability, and unnecessary full-file work. This task modernizes the existing pipeline without changing core architecture.

Scope:
- Make track selection explicit (manual > auto > whisper per primary/secondary) with deterministic reasons.
- Avoid running whisper/audio work when a track is already satisfied.
- Add bounded execution for yt-dlp and whisper subprocesses.
- Improve stage-level logging for both automatic and preprocess modes.
- Make secondary track fallback decisions explicit and not implicit.
- Preserve existing user behavior except where policy is clarified.

Files expected:
- launcher/youtube.ts
- launcher/commands/playback-command.ts (if mode/status behavior requires)
- launcher/types.ts (if schema updates needed)
- launcher/config/args-normalizer.ts (if timeout/config options added)
- launcher/util.ts (if runExternalCommand timeout controls added)
- Add/update launcher subtitle-generation tests
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Define deterministic track priority for each track: manual, then auto, then whisper (per track) and record source choice with reason.
- [ ] #2 If manual or auto satisfies a track, skip whisper for that same track and avoid unrelated full extraction/transcription work.
- [ ] #3 Introduce timeout or budget caps for yt-dlp and whisper calls; timeout should fail safe and unblock automatic playback.
- [ ] #4 Emit explicit status logs at each stage: metadata load, manual sub fetch, auto sub fetch, whisper audio extraction, whisper run, publish/load, final success/failure summary.
- [ ] #5 Make secondary handling explicit: transcribe target and translate target must only run when required by config and not by side-effect of primary logic.
- [ ] #6 Keep preprocess and automatic modes stable in success paths while making behavior in failure paths explicit and bounded.
- [ ] #7 Add tests for track-combination cases: primary available, secondary available, both missing, partial fallback, both missing with missing whisper config, timeout/error behavior.
- [ ] #8 Document any behavior changes if user-visible, especially fallback order, timeout behavior, and fallback disablement.
<!-- AC:END -->
