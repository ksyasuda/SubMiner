---
id: TASK-23.2
title: Bundle JLPT Yomitan dictionary assets for offline local lookup
status: Done
assignee: []
created_date: '2026-02-13 16:42'
updated_date: '2026-02-16 02:01'
labels: []
dependencies: []
parent_task_id: TASK-23
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Package and include the JLPT Yomitan extension dictionary assets in SubMiner so JLPT tagging can run without external network calls. Define a deterministic build/runtime path and loading strategy for the dictionary file and metadata versioning.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 JLPT dictionary asset from the existing Yomitan extension is added to the repository/build output in a tracked, offline-available location.
- [x] #2 The loader locates and opens the JLPT dictionary file deterministically at runtime.
- [x] #3 Dictionary version/source is documented so future updates are explicit and reproducible.
- [x] #4 Dictionary bundle size and load impact are documented in task notes or project docs.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Dictionary data is bundled and consumable during development and packaged app runs.
<!-- DOD:END -->
