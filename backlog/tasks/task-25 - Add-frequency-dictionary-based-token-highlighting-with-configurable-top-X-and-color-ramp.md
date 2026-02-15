---
id: TASK-25
title: >-
  Add frequency-dictionary-based token highlighting with configurable top-X and
  color ramp
status: To Do
assignee: []
created_date: '2026-02-13 16:47'
labels: []
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Leverage user-installed frequency dictionaries to color subtitle tokens based on word frequency rank, with configurable behavior: either one shared color for all words below a rank threshold or a multi-color range mapping based on frequency bands. The feature should support a configurable X (top-N words) cutoff and integrate with existing subtitle rendering flow.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Add a feature flag and configuration for frequency-based highlighting with default disabled state.
- [ ] #2 Support selecting a user-installed frequency dictionary source and reading word frequency data from it.
- [ ] #3 Introduce a configurable top-X threshold in config for which words are eligible for frequency-based coloring.
- [ ] #4 When single-color mode is enabled, all matched words within the rank rule use the configured color.
- [ ] #5 When multi-color mode is enabled, map frequency bands to colors and color tokens by their actual rank bucket.
- [ ] #6 Ensure matching is token-aware (normalization/lowercasing handling) and preserves existing subtitle tokenization behavior.
- [ ] #7 Handle missing/unsupported dictionary formats and unknown words with deterministic no-highlight fallback.
- [ ] #8 Render underline/token highlights without breaking subtitle layout or interactions.
- [ ] #9 Add tests/verification for: single-color mode, color-band mode, threshold boundary, and disabled mode.
- [ ] #10 Document dictionary source format expectations, configuration example, and performance impact of ranking lookups.
- [ ] #11 If full automatic discovery of user-installed frequency dictionaries is not possible, provide clear configuration workflow/fallback path.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Frequency-based highlighting renders using either single-color or banded-colors for valid matches, with configurable top-X threshold and documented setup.
<!-- DOD:END -->
