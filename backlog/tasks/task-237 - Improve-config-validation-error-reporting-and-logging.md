---
id: TASK-237
title: Improve config validation error reporting and logging
status: To Do
assignee: []
created_date: '2026-03-26 05:51'
labels:
  - errors
  - config
  - validation
  - ux
dependencies: []
references:
  - /docs/README.md
  - /docs/workflow/verification.md
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Replace raw error body system notifications during config validation with clearer, user-friendly summaries while retaining full technical detail in logs. The flow should surface what is wrong, where, and how to fix it without overloading the user with raw stack traces, and write structured details to console/file logs.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 When config validation fails, show a user-facing notification with cleaned, human-readable summary instead of dumping raw error text directly
- [ ] #2 Notification content includes actionable context (what field/setting failed, expected format/type, and next steps where possible)
- [ ] #3 Raw technical error details are preserved in console logs in a consistently formatted, presentable way
- [ ] #4 Config validation failures also write to persistent log file output in the same presentable format
- [ ] #5 When validation fails repeatedly or with multiple errors, aggregate and group errors for easier reading instead of showing one opaque blob
- [ ] #6 Warning/error notification should map to the specific invalid config section so users can jump to/identify what to fix
- [ ] #7 Add/update tests (unit/integration) that assert notification formatting and logging behavior for at least one malformed config case
- [ ] #8 No sensitive data (API keys/secrets) is written to logs/notifications when sanitizing errors
<!-- AC:END -->
