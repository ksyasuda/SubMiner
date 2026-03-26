---
id: TASK-236
title: Gate Jimaku and SubSync modal actions when setup is missing
status: To Do
assignee: []
created_date: '2026-03-26 05:48'
labels:
  - ui
  - setup-validation
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add safeguards in the Jimaku and SubSync modals so users cannot proceed with unsupported flows when required setup is missing. SubSync should clearly block use when alass/ffsubsync detection fails. Jimaku should surface a visible warning when no API key is configured and prevent proceeding with actions that require it.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 SubSync modal detects availability of alass and ffsubsync before enabling related action option
- [ ] #2 When only one of alass/ffsubsync is available, only the available path is selectable and clearly labeled; unavailable options are visually disabled
- [ ] #3 When neither alass nor ffsubsync are available, the unsupported action option is disabled and/or hidden, and cannot navigate to next step
- [ ] #4 If a user tries to proceed while detection says unavailable, submission is blocked with explanatory inline feedback
- [ ] #5 Jimaku modal detects missing API key (or invalid/missing key) and shows an immediate warning on search results or related UI area
- [ ] #6 When Jimaku API key is absent, actions that require key-based operations are disabled or blocked and cannot be submitted
- [ ] #7 All new/updated UX states include clear copy explaining what to fix (e.g., install binary, add API key, restart if needed)
- [ ] #8 Add or update tests for setup detection and blocked-state behavior in relevant modal/components
<!-- AC:END -->
