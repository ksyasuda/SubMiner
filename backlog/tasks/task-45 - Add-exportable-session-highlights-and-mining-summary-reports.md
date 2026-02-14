---
id: TASK-45
title: Add exportable session highlights and mining summary reports
status: To Do
assignee: []
created_date: '2026-02-14 02:17'
labels:
  - feature
  - analytics
  - immersion
  - social
dependencies:
  - TASK-28
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
After a mining session, generate a visual summary report with key metrics that can be exported as an image or markdown for sharing on language learning communities.

## Motivation
Language learners love tracking and sharing their progress. A polished session summary creates a natural sharing loop and helps users stay motivated. Communities like r/LearnJapanese, TheMoeWay Discord, and Refold regularly share immersion stats.

## Summary contents
1. **Session metrics**: Duration watched, active time, lines seen, words encountered, cards mined
2. **Efficiency stats**: Cards per hour, words per minute, lookup hit rate
3. **Top looked-up words**: The N most-looked-up words in the session
4. **Streak data**: Current daily immersion streak, total immersion hours
5. **Word cloud**: Visual word cloud of looked-up or mined vocabulary
6. **Video context**: Show/episode name, thumbnail

## Export formats
- **Image**: Rendered card-style graphic (PNG) suitable for social media sharing
- **Markdown**: Text summary for pasting into Discord/forums
- **Clipboard**: One-click copy of the summary

## Implementation notes
- Generate the image using HTML-to-canvas in the renderer, or a Node canvas library
- Data sourced from TASK-28 immersion tracking tables
- Summary could be shown automatically at session end (when mpv closes) or on demand via shortcut
- Consider a cumulative weekly/monthly summary in addition to per-session
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Session summary shows duration, lines seen, words encountered, and cards mined.
- [ ] #2 Top looked-up words are listed in the summary.
- [ ] #3 Summary is exportable as PNG image.
- [ ] #4 Summary is exportable as markdown text.
- [ ] #5 One-click copy to clipboard is supported.
- [ ] #6 Summary can be triggered on demand via shortcut or shown automatically at session end.
- [ ] #7 Data is sourced from immersion tracking (TASK-28).
<!-- AC:END -->
