---
id: TASK-104
title: Mirror overlay annotation hover behavior in vendored texthooker
status: Done
assignee:
  - codex
created_date: '2026-03-06 21:45'
updated_date: '2026-03-06 21:45'
labels:
  - texthooker
  - subtitle
  - websocket
dependencies:
  - TASK-103
references:
  - /home/sudacode/projects/japanese/SubMiner/src/core/services/subtitle-ws.ts
  - /home/sudacode/projects/japanese/SubMiner/vendor/texthooker-ui/src/components/App.svelte
  - /home/sudacode/projects/japanese/SubMiner/vendor/texthooker-ui/src/line-markup.ts
  - /home/sudacode/projects/japanese/SubMiner/vendor/texthooker-ui/src/app.css
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Bring bundled texthooker annotation rendering closer to the visible overlay. Keep the lightweight texthooker UX, but preserve token metadata for hover, match overlay color-precedence rules across known/N+1/name/frequency/JLPT, expose name-match highlighting as a toggle, and emit a structured annotation payload on the dedicated websocket so non-SubMiner clients can treat it as an API.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [x] #1 Annotation websocket payload includes both rendered `sentence` HTML and structured token metadata for generic clients.
- [x] #2 Vendored texthooker preserves annotation metadata attrs needed for hover labels and uses overlay-matching color precedence rules.
- [x] #3 Vendored texthooker supports character-name highlighting with a user-facing toggle and standalone-web note.
- [x] #4 Hovering annotated texthooker tokens reveals JLPT/frequency metadata without adding the full overlay popup workflow.
- [x] #5 Focused serializer, texthooker markup, socket parsing, CSS, and build verification pass.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Extended the dedicated annotation websocket payload to ship `version`, plain `text`, rendered `sentence`, and structured `tokens` metadata while keeping backward-compatible `sentence` consumers working. Updated the vendored texthooker to preserve hover metadata attrs, follow overlay color precedence for known/N+1/name/frequency/JLPT annotations, add a character-name highlight toggle plus standalone-web dictionary note, and render lightweight hover labels for frequency/JLPT metadata. Added focused regression coverage and rebuilt both the vendored texthooker bundle and SubMiner.
<!-- SECTION:FINAL_SUMMARY:END -->
