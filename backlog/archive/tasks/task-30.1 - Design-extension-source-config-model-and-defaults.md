---
id: TASK-30.1
title: Design extension source config model and defaults
status: To Do
assignee: []
created_date: '2026-02-13 18:32'
updated_date: '2026-02-13 18:34'
labels: []
dependencies: []
parent_task_id: TASK-30
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Define a backend-agnostic configuration contract for extension repository streaming, including resolver endpoints/process mode, query/auth headers, timeouts, enable flags, and source preference. Wire schema through Config/ResolvedConfig and generated template/defaults so users can manage repos entirely through config.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Add new config sections for extension source providers in Config and ResolvedConfig types.
- [ ] #2 Add validation defaults and env-compatible parsing for provider list, auth, header overrides, and feature flags.
- [ ] #3 Update config template and docs text so defaults are discoverable and editable.
- [ ] #4 Invalid/missing config should fail fast with a clear message path.
- [ ] #5 Existing config readers do not regress when no providers are configured.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Phase 1 — Foundation: config contract + validation + defaults

<!-- SECTION:NOTES:END -->

## Definition of Done

<!-- DOD:BEGIN -->

- [ ] #1 Config examples in template/docs include at least one provider entry shape.
- [ ] #2 Defaults remain backward-compatible when key is absent.
- [ ] #3 Feature can be disabled without touching unrelated settings.
<!-- DOD:END -->
