---
id: TASK-25
title: >-
  Add frequency-dictionary-based token highlighting with configurable top-X and
  color ramp
status: Done
assignee: []
created_date: '2026-02-13 16:47'
updated_date: '2026-02-16 06:48'
labels: []
dependencies: []
documentation:
  - /Users/sudacode/.codex/worktrees/2089/SubMiner/docs/configuration.md
  - /Users/sudacode/.codex/worktrees/2089/SubMiner/docs/jlpt-vocab-bundle.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Leverage user-installed frequency dictionaries to color subtitle tokens based on word frequency rank, with configurable behavior: either one shared color for all words below a rank threshold or a multi-color range mapping based on frequency bands. The feature should support a configurable X (top-N words) cutoff and integrate with existing subtitle rendering flow.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Add a feature flag and configuration for frequency-based highlighting with default disabled state.
- [x] #2 Support selecting a user-installed frequency dictionary source and reading word frequency data from it.
- [x] #3 Introduce a configurable top-X threshold in config for which words are eligible for frequency-based coloring.
- [x] #4 When single-color mode is enabled, all matched words within the rank rule use the configured color.
- [x] #5 When multi-color mode is enabled, map frequency bands to colors and color tokens by their actual rank bucket.
- [x] #6 Ensure matching is token-aware (normalization/lowercasing handling) and preserves existing subtitle tokenization behavior.
- [x] #7 Handle missing/unsupported dictionary formats and unknown words with deterministic no-highlight fallback.
- [x] #8 Render underline/token highlights without breaking subtitle layout or interactions.
- [x] #9 Add tests/verification for: single-color mode, color-band mode, threshold boundary, and disabled mode.
- [x] #10 Document dictionary source format expectations, configuration example, and performance impact of ranking lookups.
- [x] #11 If full automatic discovery of user-installed frequency dictionaries is not possible, provide clear configuration workflow/fallback path.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-02-16: Updated docs for frequency dictionary behavior. Clarified built-in  fallback,  precedence, and shared  format expectations in  and .

Added docs references for frequency dictionary defaults and fallback behavior.

As of 2026-02-16, docs and implementation are considered complete for TASK-25; frequency highlighting fallback, custom sourcePath precedence, topX, single/banded modes, token pipeline integration, and fallback behavior are present; documentation and tests exist in src/core/services and src/renderer.

2026-02-16: Frequency-dictionary highlighting feature fully complete and shipped. Task acceptance criteria, DoD, and docs alignment are all marked complete in this task record.
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Frequency-based highlighting renders using either single-color or banded-colors for valid matches, with configurable top-X threshold and documented setup.
<!-- DOD:END -->
