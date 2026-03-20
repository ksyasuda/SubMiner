---
id: TASK-125
title: Add native AI API key secret storage
status: To Do
assignee: []
created_date: '2026-03-08 07:25'
updated_date: '2026-03-18 05:27'
labels:
  - ai
  - config
  - security
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/src/ai/client.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/anilist/anilist-token-store.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/jellyfin-token-store.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/main.ts
priority: medium
ordinal: 2000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Store the shared AI provider API key using the app's native secret-storage pattern so users do not need to keep the OpenRouter key in config files or shell commands.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Users can configure the shared AI provider without storing the API key in config.jsonc.
- [ ] #2 The app persists and reloads the shared AI API key using encrypted native secret storage when available.
- [ ] #3 Behavior is defined for existing ai.apiKey and ai.apiKeyCommand configs, including compatibility during migration.
- [ ] #4 The feature has regression tests covering key resolution and storage behavior.
- [ ] #5 User-facing configuration/docs are updated to describe the supported setup.
<!-- AC:END -->
