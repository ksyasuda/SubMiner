---
id: TASK-23
title: >-
  Add opt-in JLPT level tagging by bundling and querying local Yomitan
  dictionary
status: Done
assignee: []
created_date: '2026-02-13 16:42'
updated_date: '2026-02-18 04:11'
labels: []
dependencies: []
priority: high
ordinal: 34000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Implement an opt-in JLPT token annotation feature that annotates subtitle words with JLPT level in-session. The feature should use a bundled dictionary source from the existing JLPT Yomitan extension, parse/query its dictionary file to determine whether a token appears and its JLPT level, and render token-level visual tags as a colored underline spanning each token length. Colors must correspond to JLPT levels (e.g., N5/N4/N3/N2/N1) using a consistent mapping.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Add an opt-in setting/feature flag so JLPT tagging is disabled by default and can be enabled per user/session as requested.
- [x] #2 Bundle the existing JLPT Yomitan extension package/data into the project so lookups can be performed offline from local files.
- [x] #3 Implement token-level dictionary lookup against the bundled JLPT dictionary file to determine presence and JLPT level for words in subtitle lines.
- [x] #4 Render a colored underline under each token determined to have a JLPT level; the underline must match token width/length and not affect layout or disrupt line rendering.
- [x] #5 Assign different underline colors per JLPT level (at minimum N5/N4/N3/N2/N1) with a stable mapping documented in task notes.
- [x] #6 Handle unknown/no-match tokens as non-tagged while preserving existing subtitle styling and interaction behavior.
- [x] #7 When disabled, no JLPT lookups are performed and subtitles render exactly as current behavior.
- [x] #8 Add tests or deterministic checks covering at least one positive match, one non-match, and one unknown/unsupported-level fallback path.
- [x] #9 Document expected dictionary source and any size/performance impact of bundling the JLPT extension data.
- [x] #10 If dictionary format/version constraints block exact level extraction, the task includes explicit limitation notes and a deterministic fallback strategy.
<!-- AC:END -->

## Definition of Done

<!-- DOD:BEGIN -->

- [x] #1 Feature has a clear toggle and persistence of preference if applicable.
- [x] #2 JLPT rendering is visually verified for all supported levels with distinct colors and no overlap/regression in subtitle legibility.
<!-- DOD:END -->
