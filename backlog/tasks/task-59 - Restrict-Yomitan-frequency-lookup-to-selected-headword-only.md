---
id: TASK-59
title: Restrict Yomitan frequency lookup to selected headword only
status: Done
assignee: []
created_date: '2026-02-16 22:16'
updated_date: '2026-02-16 22:18'
labels: []
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Update tokenizer and related scripts/tests so frequency lookup no longer uses Yomitan headword variant lists and instead only uses the selected headword returned by Yomitan.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Frequency ranking for Yomitan tokens uses only the token headword (with existing fallback behavior) and not `frequencyLookupTerms` variants.
- [x] #2 Tokenizer tests reflect the new headword-only lookup behavior.
- [x] #3 Parser testing script output no longer implies variant-based frequency lookup.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Updated frequency lookup to use only the selected token lookup text (headword first, fallback to reading/surface only when headword is absent) and removed Yomitan variant-term usage. Removed `frequencyLookupTerms` from token mapping/types, updated tokenizer tests for headword-only behavior, and aligned helper scripts (`scripts/get_frequency.ts`, `scripts/test-yomitan-parser.ts`) so diagnostics/output no longer imply variant-based lookup.
<!-- SECTION:FINAL_SUMMARY:END -->
