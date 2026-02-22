# TASK-100 Dead Code Report (2026-02-22)

## Baseline Verification

- `bun run build` -> PASS
- `bun run test:fast` -> PASS

## Discovery Commands

- `tsc --noEmit --noUnusedLocals --noUnusedParameters`
- `bunx ts-prune -p tsconfig.json`

## Triage

### Remove

- `src/anki-connect.ts` - removed unused `url` instance field.
- `src/anki-integration.ts` - removed unused wrappers: `poll`, `showProgressTick`, `refreshMiscInfoField`.
- `src/anki-integration/card-creation.ts` - removed unused `MediaGenerator` import.
- `src/anki-integration/ui-feedback.ts` - removed unused callback parameter in `withUpdateProgress`.
- `src/core/services/anki-jimaku-ipc.ts` - removed unused `JimakuDownloadQuery` import.
- `src/core/services/immersion-tracker-service.ts` - removed unused fields `lastMaintenanceMs`, `lastQueueWriteAtMs`; removed unused `runRollupMaintenance` wrapper.
- `src/core/services/ipc-command.ts` - removed unused `RuntimeOptionValue` import.
- `src/renderer/positioning/position-state.ts` - removed unused `ctx` parameter from `getPersistedOffset`.
- `src/tokenizers/index.ts` - removed unused exported helpers `getRegisteredTokenizerProviderIds`, `createTokenizerProvider`.
- `src/token-mergers/index.ts` - removed unused exported helpers `getRegisteredTokenMergerProviderIds`, `createTokenMergerProvider`.
- `src/core/utils/index.ts` - removed unused barrel re-exports `asBoolean`, `asFiniteNumber`, `asString`.

### Keep (intentional / out-of-scope)

- `src/main/runtime/composers/composer-contracts.type-test.ts` private `_` type aliases remain; they are compile-time contract assertions.
- `src/main.ts` large unused-import cluster from ongoing composer/runtime decomposition kept for separate focused task to avoid behavior risk.
- Broad `ts-prune` type-export findings in `src/types.ts` and multiple domain modules kept; many are declaration-surface exports and module-local false positives.

## Complexity Delta

- Removed 13 confirmed dead declarations/imports/helpers.
- Removed 4 unused exported entrypoints from provider registries/util barrel.
- `tsc --noEmit --noUnusedLocals --noUnusedParameters` diagnostics reduced to `39` lines; remaining diagnostics are concentrated in `src/main.ts` plus intentional type-test aliases.

## Regression Safety / Tests

- `bun test src/anki-integration.test.ts src/core/services/mining.test.ts src/core/services/anki-jimaku-ipc.test.ts src/core/services/immersion-tracker-service.test.ts src/core/services/ipc.test.ts`
  - partial pass; direct IPC test invocation hit Electron ESM test harness issue (`Export named 'ipcMain' not found`) unrelated to cleanup.
- Required task gates:
  - `bun run build` -> PASS
  - `bun run test:core:src` -> PASS
  - `bun run test:config:src` -> PASS
  - `bun run check:file-budgets` -> PASS (warning mode, no strict hotspot violations)

## Remaining Candidates

- Continue with dedicated `src/main.ts` dead-import cleanup once runtime composer migration settles.
- Revisit `ts-prune` findings with a declaration-aware filter to separate true dead exports from public API type surfaces.
