---
id: TASK-99
title: Add configurable character dictionary collapsible section open states
status: Done
assignee: []
created_date: '2026-03-07 00:00'
updated_date: '2026-03-07 00:00'
labels:
  - dictionary
  - config
references:
  - /home/sudacode/projects/japanese/SubMiner/src/main/character-dictionary-runtime.ts
  - /home/sudacode/projects/japanese/SubMiner/src/config/resolve/integrations.ts
  - /home/sudacode/projects/japanese/SubMiner/src/config/definitions/defaults-integrations.ts
priority: medium
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add per-section config for character dictionary collapsible glossary sections so Description, Character Information, and Voiced by can each default open or closed independently. Default all sections closed.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Config supports `anilist.characterDictionary.collapsibleSections.description`.
- [x] #2 Config supports `anilist.characterDictionary.collapsibleSections.characterInformation`.
- [x] #3 Config supports `anilist.characterDictionary.collapsibleSections.voicedBy`.
- [x] #4 Default config keeps all generated character dictionary collapsible sections closed.
- [x] #5 Regression coverage verifies config parsing/warnings and generated glossary `details.open` behavior.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added per-section open-state config under `anilist.characterDictionary.collapsibleSections` for `description`, `characterInformation`, and `voicedBy`, all defaulting to `false`. Wired the glossary generator to read those settings so generated `details.open` matches config, and added regression coverage for defaults, parsing/warnings, registry exposure, and runtime glossary output.
<!-- SECTION:FINAL_SUMMARY:END -->
