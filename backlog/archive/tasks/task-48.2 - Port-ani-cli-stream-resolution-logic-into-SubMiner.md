---
id: TASK-48.2
title: Port ani-cli stream-resolution logic into SubMiner
status: To Do
assignee: []
created_date: '2026-02-14 06:03'
updated_date: '2026-02-14 08:19'
labels:
  - stream
  - ani-cli
dependencies: []
parent_task_id: TASK-48
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Implement stream URL resolution by porting ani-cli logic for selecting providers/episodes and obtaining playable stream URLs so SubMiner can consume stream sources directly.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Encapsulate stream search/provider selection logic in a dedicated module in SubMiner.
- [ ] #2 Resolve episode query input into a canonical playable stream URL in streaming mode.
- [ ] #3 Preserve existing behavior for non-streaming flow and expose errors when stream resolution fails.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Superseded by TASK-51. Stream URL resolution via ani-cli postponed; previous attempt exposed anti-bot/403 fragility and poor title-source reliability.
<!-- SECTION:NOTES:END -->
