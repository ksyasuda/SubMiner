---
id: TASK-41
title: Add real-time sentence difficulty scoring with auto-pause on hard lines
status: To Do
assignee: []
created_date: '2026-02-14 02:09'
labels:
  - feature
  - nlp
  - immersion
  - subtitle
dependencies:
  - TASK-23
  - TASK-25
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Compute a real-time difficulty score for each subtitle line using JLPT level data (TASK-23) and frequency dictionary data (TASK-25), and use this score to drive smart playback features.

## Motivation
Learners at different levels have different needs. N4 learners want to pause on N2+ lines; advanced learners want to skip easy content. A per-line difficulty score enables intelligent playback that adapts to the learner's level.

## Features
1. **Per-line difficulty score**: Combine JLPT levels and frequency ranks of tokens to produce a composite difficulty score (e.g., 1-5 scale or JLPT-equivalent label)
2. **Visual difficulty indicator**: Subtle color/icon on each subtitle line indicating difficulty
3. **Auto-pause on difficult lines**: Configurable threshold — pause playback when a line exceeds the user's set difficulty level
4. **Per-episode difficulty rating**: Average difficulty across all lines, shown in the episode browser (TASK-34)
5. **Difficulty trend within a video**: Show whether difficulty increases/decreases over the episode (useful for detecting climax scenes with complex dialogue)

## Scoring algorithm (suggested)
- For each token in a line, look up JLPT level (N5=1, N1=5) and frequency rank
- Weight unknown words (not in Anki known-word cache from TASK-24) more heavily
- Composite score = weighted average of token difficulties, with bonus for line length and grammar complexity
- Configurable weights so users can tune sensitivity

## Design constraints
- Scoring must run synchronously during subtitle rendering without perceptible latency
- Score computation should be cached per subtitle line (lines repeat on seeks/replays)
- Auto-pause should be debounced to avoid rapid pause/unpause on sequential hard lines
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Each subtitle line receives a difficulty score based on JLPT and frequency data.
- [ ] #2 A visual indicator shows per-line difficulty in the overlay.
- [ ] #3 Auto-pause triggers when a line exceeds the user's configured difficulty threshold.
- [ ] #4 Difficulty scoring does not add perceptible latency to subtitle rendering.
- [ ] #5 Per-episode average difficulty is available for the episode browser.
- [ ] #6 Scoring weights are configurable in settings.
<!-- AC:END -->
