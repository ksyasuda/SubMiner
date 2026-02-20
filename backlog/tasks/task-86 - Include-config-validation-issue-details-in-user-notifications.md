---
id: TASK-86
title: Include config validation issue details in user notifications
status: Done
assignee: []
created_date: '2026-02-19 17:24'
updated_date: '2026-02-19 23:18'
labels: []
dependencies: []
priority: medium
ordinal: 62000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
When config validation finds non-fatal issues, users should see concise per-issue details in notification body (not only issue count) so they can fix config quickly without checking logs.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Startup notification body includes per-issue details (path + message) for config validation warnings.
- [x] #2 Hot-reload validation warning notifications include per-issue details in notification body.
- [x] #3 Notification text remains concise and does not exceed practical desktop notification limits.
- [x] #4 Automated tests cover notification body formatting with detailed issues.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added `buildConfigWarningNotificationBody` to format concise multi-line warning details (path+message, line limit + overflow count). Startup warnings now use this formatter for desktop notification body. Config hot-reload runtime now emits non-fatal validation warnings via `onValidationWarnings(configPath, warnings)` and main process surfaces them through desktop notifications. Added tests for formatter output, startup notification body content, and hot-reload warning callback behavior.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Config validation notifications now include concrete issue details in the notification body instead of only a count. Startup and hot-reload warning paths both surface per-issue `path: message` lines with concise truncation safeguards. Added regression tests covering formatter output and both notification paths.
<!-- SECTION:FINAL_SUMMARY:END -->
