---
id: TASK-46
title: >-
  Add listening-only mode with on-demand subtitle reveal for comprehension
  training
status: To Do
assignee: []
created_date: '2026-02-14 02:19'
labels:
  - feature
  - immersion
  - subtitle
  - listening
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Implement a listening-only mode that hides subtitles by default and only reveals them on demand (click, tap, or shortcut), tracking how many lines the user understood without looking.

## Motivation
Subtitle dependency is a common trap for language learners. Users watch hundreds of hours with subtitles but struggle when subtitles are removed. A listening-only mode provides a structured way to wean off subtitles: watch without them, check when needed, and track comprehension over time.

## Features
1. **Hidden-by-default subtitles**: Subtitles are loaded and tracked but not displayed
2. **On-demand reveal**: Press a key or click to reveal the current subtitle line briefly (3-5 seconds, then auto-hides)
3. **Comprehension tracking**: Track reveal rate (lines revealed / total lines) as a listening comprehension metric
4. **Difficulty-aware reveal**: Optionally auto-reveal lines above a difficulty threshold (pairs with sentence difficulty scoring feature)
5. **Session stats**: At session end, show listening comprehension percentage
6. **Progressive mode**: Start with subtitles visible, then fade them out after N minutes (configurable ramp)

## Technical considerations
- Subtitle timing is still tracked (for Anki card creation if user mines a revealed line)
- MPV's native subtitles should be hidden; SubMiner handles visibility
- Reveal animation should be smooth (fade in/out, not jarring)
- Mining workflow should still work: reveal → click word → Yomitan → Anki
- Comprehension data feeds into TASK-28 immersion tracking if available

## Design constraints
- Must not break existing subtitle display modes
- Listening mode should be a toggle (shortcut to enter/exit)
- Reveal count and comprehension rate should persist with session data
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Listening mode hides subtitles by default while continuing to track timing.
- [ ] #2 Shortcut or click reveals the current subtitle briefly (configurable duration).
- [ ] #3 Reveal rate (lines revealed / total lines) is tracked per session.
- [ ] #4 Mining workflow works on revealed lines (click word → Yomitan → Anki).
- [ ] #5 Session end shows listening comprehension percentage.
- [ ] #6 Mode is toggleable via shortcut without restarting.
- [ ] #7 Subtitle reveal has smooth fade in/out animation.
<!-- AC:END -->
