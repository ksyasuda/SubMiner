---
id: TASK-331
title: Address unresolved CodeRabbit comments on PR 57
status: Done
assignee:
  - codex
created_date: '2026-05-04 03:21'
updated_date: '2026-05-04 03:27'
labels:
  - pr-feedback
  - coderabbit
dependencies: []
references:
  - 'https://github.com/ksyasuda/SubMiner/pull/57'
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Assess and fix unresolved CodeRabbit review comments on PR #57 after rebasing tokenizer-updates. Scope includes manual clipboard SentenceAudio guard, tokenizer standalone particle blacklist, AniList guessit fallback confidence, startup gate duplicate auto-start, and small regression-test hardening where applicable.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Each unresolved CodeRabbit comment is either fixed or explicitly assessed as not applicable against current code.
- [x] #2 Regression tests cover behavior changes where practical.
- [x] #3 Relevant focused tests and typecheck pass.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Fixed all verified actionable CodeRabbit comments from PR #57: manual clipboard updates no longer fall back to ExpressionAudio when SentenceAudio is absent, connective particle phrases no longer suppress lexical verb readings like 立って, guessit output only borrows parser season/episode from non-low-confidence parses, duplicate auto-start no longer releases an active pause-until-ready gate, JLPT CSS tests block text-decoration shorthand underlines, post-watch update rejection logging is covered, and duplicate quit-on-disconnect predicate code is shared.

Verification: bun test src/anki-integration/card-creation-manual-update.test.ts src/core/services/tokenizer/annotation-stage.test.ts src/core/services/anilist/anilist-updater.test.ts src/main/runtime/mpv-main-event-actions.test.ts src/renderer/subtitle-render.test.ts; lua scripts/test-plugin-start-gate.lua; bun run typecheck; bun run test:fast.
<!-- SECTION:NOTES:END -->
