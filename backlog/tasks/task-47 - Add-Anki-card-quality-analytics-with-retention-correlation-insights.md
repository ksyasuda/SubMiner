---
id: TASK-47
title: Add Anki card quality analytics with retention correlation insights
status: To Do
assignee: []
created_date: '2026-02-14 02:21'
labels:
  - feature
  - anki
  - analytics
  - immersion
dependencies:
  - TASK-28
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Analyze existing Anki cards created by SubMiner to identify which card characteristics correlate with better retention, helping users understand what makes a "good" mining card for them personally.

## Motivation

Not all mined cards are equal. Some are remembered easily; others become leeches. By analyzing retention data from Anki alongside card characteristics (sentence length, word frequency, JLPT level, context richness), SubMiner can provide personalized insights about optimal mining strategies.

## Features

1. **Card retention analysis**: Query Anki for review history of SubMiner-created cards, compute retention rates
2. **Characteristic correlation**: Correlate retention with:
   - Sentence length (words/characters)
   - Target word frequency rank
   - Target word JLPT level
   - Number of unknown words in the sentence (i+N analysis)
   - Whether audio/screenshot was included
   - Source media genre/type
3. **Insights dashboard**: Show actionable insights like "Your best-retained cards have 8-15 words and 1-2 unknown words"
4. **Mining recommendations**: Real-time suggestions during mining — "This sentence has 4 unknown words; consider a simpler example"
5. **Leech prediction**: Flag newly created cards that match the profile of past leeches

## Technical considerations

- Requires querying Anki review history via AnkiConnect (cardInfo, getReviewsOfCards)
- Analysis can run as a background task during idle time
- Results should be cached locally (SQLite via TASK-28 or separate store)
- Privacy-sensitive: all analysis is local, no data leaves the machine
- Consider batch analysis (run nightly or on demand) vs real-time

## Design constraints

- Must not slow down the mining workflow
- Insights should be actionable, not just statistical
- Analysis should work with existing card format (no retroactive changes needed)
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Retention rates are computed for SubMiner-created Anki cards via AnkiConnect.
- [ ] #2 Correlations between card characteristics and retention are computed and displayed.
- [ ] #3 Insights dashboard shows actionable recommendations (optimal sentence length, i+N target).
- [ ] #4 Real-time mining suggestions appear when creating cards with suboptimal characteristics.
- [ ] #5 Analysis runs in background without impacting mining performance.
- [ ] #6 All analysis is local — no data sent externally.
<!-- AC:END -->
