---
id: TASK-42
title: Add spaced repetition video clip review mode for mined Anki cards
status: To Do
assignee: []
created_date: '2026-02-14 02:11'
labels:
  - feature
  - anki
  - immersion
  - review
dependencies: []
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Implement a "review mode" that replays short video clips from previously mined Anki cards, allowing users to review vocabulary in its original immersion context rather than as isolated text cards.

## Motivation
Traditional Anki review shows a sentence on a card. But the original context — hearing the word spoken, seeing the scene, feeling the emotion — is far more memorable. SubMiner already captures audio clips and screenshots for cards. This feature uses that data to create a video-based review experience.

## Workflow
1. User activates review mode (shortcut or menu)
2. SubMiner queries AnkiConnect for due/new cards from a configured deck
3. For each card, SubMiner locates the source video and timestamp (stored in card metadata)
4. Plays a 3-8 second clip around the mined sentence with subtitles hidden
5. User tries to recall the target word/meaning
6. On reveal: show the subtitle, highlight the target word, display the Anki card fields
7. User grades the card (easy/good/hard/again) via shortcuts, fed back to Anki via AnkiConnect

## Technical considerations
- Requires source video files to still be accessible (local path stored in card)
- Falls back gracefully if video is unavailable (show audio clip + screenshot instead)
- MPV can seek to timestamp and play a range (`--start`, `--end` flags or `loadfile` with options)
- Card metadata needs a consistent format for source video path and timestamp (may need to standardize the field SubMiner writes)
- Consider a playlist mode that queues multiple clips for batch review

## Design constraints
- Must work with existing Anki card format (no retroactive card changes required)
- Graceful degradation when source video is missing
- Review session should be pausable and resumable
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Review mode queries AnkiConnect for due cards and locates source video + timestamp.
- [ ] #2 Video clips play with subtitles hidden, then revealed on user action.
- [ ] #3 Target word is highlighted in the revealed subtitle.
- [ ] #4 User can grade cards (easy/good/hard/again) via keyboard shortcuts.
- [ ] #5 Grades are sent back to Anki via AnkiConnect.
- [ ] #6 Graceful fallback when source video file is unavailable (audio + screenshot).
- [ ] #7 Review session can be paused and resumed.
<!-- AC:END -->
