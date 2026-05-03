---
id: TASK-317
title: Add browser open affordance for texthooker
status: Done
assignee: []
created_date: '2026-05-03 02:02'
updated_date: '2026-05-03 02:21'
labels:
  - feature
  - texthooker
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add a `-o` flag to the texthooker subcommand to open the texthooker page in the user's default browser, and add a tray app option that triggers the same behavior. Implement with tests and existing launcher/tray patterns.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `texthooker -o` starts/targets the texthooker page and opens it in the default browser.
- [x] #2 Tray app exposes a menu option to open the texthooker page in the default browser.
- [x] #3 Existing texthooker behavior without `-o` remains unchanged.
- [x] #4 Relevant CLI/tray behavior covered by tests.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented `subminer texthooker -o` by parsing the launcher subcommand flag, forwarding `--open-browser` to the app texthooker command, and allowing that app arg to force browser opening even when `texthooker.openBrowser` is false. Added an `Open Texthooker` tray menu item wired through the same CLI command path. Updated docs-site usage/launcher/API docs and added a changelog fragment. Verification: targeted CLI/tray tests passed; `bun run typecheck` passed; `bun run docs:test` passed; `bun run changelog:lint` passed; `bun run test:env` passed; `bun run build` passed; `bun run test:smoke:dist` passed; `bun run docs:build` passed after installing docs-site deps. `bun run test:fast` is blocked by an existing broader-suite failure in `runSubsyncManual writes deterministic _retimed filename when replace is false` (`window.electronAPI` undefined), followed by Bun nested-test cascade errors.

Follow-up fix: `subminer texthooker -o` now opens `http://127.0.0.1:5174` from the launcher after a successful texthooker app handoff, so it works even when the installed SubMiner app binary does not yet understand the app-side `--open-browser` flag. Reproduced the reported behavior; confirmed the texthooker server was running at `127.0.0.1:5174`; added a launcher regression asserting the browser URL is opened. Verification: `bun test launcher/mpv.test.ts launcher/config/cli-parser-builder.test.ts launcher/config/args-normalizer.test.ts src/core/services/cli-command.test.ts src/main/runtime/tray-runtime.test.ts src/main/runtime/tray-main-actions.test.ts` passed; `bun run typecheck` passed; `bun run build:launcher` passed.
<!-- SECTION:FINAL_SUMMARY:END -->
