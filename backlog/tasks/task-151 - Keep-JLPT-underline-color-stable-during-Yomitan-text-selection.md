---
id: TASK-151
title: Keep JLPT underline color stable during Yomitan text selection
status: Done
assignee:
  - OpenCode
created_date: '2026-03-10 06:42'
updated_date: '2026-03-10 07:54'
labels: []
dependencies: []
references:
  - src/renderer/style.css
  - src/renderer/subtitle-render.test.ts
documentation:
  - ../subminer-docs/development.md
  - ../subminer-docs/architecture.md
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Fix the subtitle overlay so JLPT underlines keep their JLPT color when Yomitan lookups trigger hover/selection styling on tokens that are also known, N+1, or frequency-highlighted.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 JLPT-tagged subtitle tokens keep their JLPT underline color during hover/selection states used by Yomitan lookup.
- [x] #2 Tokens that combine JLPT with known, N+1, name-match, or frequency styling preserve their primary text color while keeping the JLPT underline color.
- [x] #3 Renderer regression coverage verifies the JLPT underline color remains explicit in the stylesheet for combined styling states.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Revert the temporary box-shadow JLPT marker experiment and restore native JLPT underline styling in `src/renderer/style.css`.
2. Add explicit JLPT hover/selection color rules with higher specificity so Yomitan lookup highlight states cannot repaint the underline to the token hover color.
3. Update `src/renderer/subtitle-render.test.ts` to verify JLPT hover/selection rules keep the predefined JLPT underline color, then rerun focused and relevant verification commands.

4. Add JLPT-specific `text-decoration-color` overrides directly on child `.c` spans and their hover/selection states, since Yomitan lookup likely paints the decorated child text runs rather than only the parent token span.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added a renderer regression test that asserts JLPT underline color stays explicit through hover and selection styling.

Updated `src/renderer/style.css` to route JLPT underline color through `--subtitle-jlpt-underline-color` and re-apply it on word/character hover and selection states.

Verification: `bun test src/renderer/subtitle-render.test.ts`, `bun run typecheck`, `bun run test:fast`, `bun run build`, `bun run changelog:lint`, `bunx prettier --check src/renderer/style.css src/renderer/subtitle-render.test.ts changes/jlpt-underline-yomitan.md`. `bun run format:check:src` still fails because of unrelated existing formatting issues in `src/main-entry-runtime.test.ts` and `src/main-entry-runtime.ts`.

User verified the original CSS variable propagation fix did not resolve the live issue; continuing investigation and reopening the task.

Confirmed Chromium was repainting native `text-decoration` underlines with the selected text color; reproduced it in a temporary browser repro and verified that switching the JLPT marker to an inset box-shadow keeps the JLPT color stable during selection.

Replaced the previous native underline approach with a fragment-safe JLPT marker using `box-shadow` plus `box-decoration-break`, then updated stylesheet regression coverage to match the new implementation.

Verification: `bun test src/renderer/subtitle-render.test.ts`, `bunx prettier --check src/renderer/style.css src/renderer/subtitle-render.test.ts`, `bun run typecheck`, `bun run test:fast`, `bun run build`.

User requested reverting the box-shadow marker experiment. Narrowing the fix to preserve native JLPT underline color during Yomitan hover/selection instead of changing the underline rendering method.

Reverted the box-shadow experiment per user request. Restored native JLPT underline styling and added explicit JLPT hover/selection rules so Yomitan lookup states keep the predefined JLPT underline color with `!important`.

Verification after revert/fix: `bun test src/renderer/subtitle-render.test.ts`, `bunx prettier --check src/renderer/style.css src/renderer/subtitle-render.test.ts`, `bun run typecheck`, `bun run build`.

Issue still reproduces specifically during Shift-held Yomitan lookup. Investigating decoration propagation onto child `.c` spans during lookup-triggered hover/selection states.

Added JLPT-specific `text-decoration-color` overrides directly to child `.c` spans and their hover/selection states, in addition to the token span itself, so lookup-triggered decoration painting on child runs keeps the JLPT color.

Verification: `bun test src/renderer/subtitle-render.test.ts`, `bunx prettier --check src/renderer/style.css src/renderer/subtitle-render.test.ts`, `bun run typecheck`, `bun run build`.

Added a final high-specificity combined-state JLPT override after the annotation hover/selection rules. It targets tokens that are both JLPT-tagged and known/N+1/name-match/frequency-highlighted, forcing `text-decoration-color` back to the JLPT variable after the later annotation state blocks.

Verification: `bun test src/renderer/subtitle-render.test.ts`, `bunx prettier --check src/renderer/style.css src/renderer/subtitle-render.test.ts`, `bun run build`.

Switched the JLPT underline from native `text-decoration` to a separate painted underline layer using `background-image` on JLPT tokens. Also changed hover token highlight rules to use `background-color` so the JLPT underline layer survives Shift+Yomitan hover/lookup interactions.

Verification: `bun test src/renderer/subtitle-render.test.ts`, `bunx prettier --check src/renderer/style.css src/renderer/subtitle-render.test.ts`, `bun run build`.

User reports the latest background-image JLPT underline change introduced grey-looking token highlights. Investigating the interaction between the new underline layer and hover background styling.

The background-image underline layer interacts poorly with hover fills and makes highlighted tokens look grey. Switching the separate JLPT underline layer to an inset box-shadow so hover background color and JLPT underline paint can coexist cleanly.

Adjusted the separate JLPT underline layer from `background-image` to inset `box-shadow`. This keeps the underline decoupled from Yomitan hover color while avoiding the grey-looking underline artifact introduced by the background-image approach.

User reports the separate underline layer still disappears/greys out during lookup. Switching strategy again: use a dedicated `border-bottom` JLPT marker instead of box-shadow/background so the underline stays independent from Yomitan hover paint without blending artifacts.

Replaced the separate JLPT underline layer with `border-bottom` plus a small bottom padding. This keeps the underline visually separate from Yomitan hover paint and avoids the disappearing/grey artifact seen with background-image and box-shadow.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Switched the JLPT underline layer in `src/renderer/style.css` to `border-bottom` plus `padding-bottom`, keeping it independent from Yomitan hover/selection repainting while avoiding the disappearing or grey artifact seen with the background-image and box-shadow approaches.

Updated `src/renderer/subtitle-render.test.ts` to assert the border-bottom-based JLPT underline layer.

Verification: `bun test src/renderer/subtitle-render.test.ts`, `bunx prettier --check src/renderer/style.css src/renderer/subtitle-render.test.ts`, `bun run build`.
<!-- SECTION:FINAL_SUMMARY:END -->
