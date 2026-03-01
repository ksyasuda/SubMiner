---
id: TASK-76
title: 'Tokenizer: remove POS exclusion config surface and keep hardcoded defaults'
status: Done
assignee: []
created_date: '2026-03-01 02:45'
updated_date: '2026-03-01 04:14'
labels: []
dependencies: []
priority: medium
ordinal: 5000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Remove user-facing config keys for annotation POS exclusions. Keep N+1/frequency POS exclusion behavior as built-in defaults with no config required.

Scope: remove config parsing/registry/docs/example for annotationFilters.pos1Exclusions/pos2Exclusions while preserving runtime filtering behavior.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 No user-facing config option exists for annotation POS exclusions.
- [x] #2 Runtime N+1/frequency exclusion behavior remains active via built-in defaults.
- [x] #3 Config/docs/example/tests updated accordingly.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Removed user-facing subtitleStyle.annotationFilters POS exclusion configuration (schema/resolver/options/docs/example). POS-based N+1/frequency filtering now always uses built-in defaults in runtime. Preserved robust exclusion behavior including merged-token overlap POS handling and N+1-only MeCab enrichment path.

<!-- SECTION:FINAL_SUMMARY:END -->
