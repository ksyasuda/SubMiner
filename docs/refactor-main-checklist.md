# Main.ts Refactor Checklist

This checklist is the safety net for `src/main.ts` decomposition work.

## Invariants (Do Not Break)

- Keep all existing CLI flags and aliases working.
- Keep IPC channel names and payload shapes backward-compatible.
- Preserve overlay behavior:
  - visible overlay toggles and follows expected state
  - invisible overlay toggles and mouse passthrough behavior
- Preserve MPV integration behavior:
  - connect/disconnect flows
  - subtitle updates and overlay updates
- Preserve texthooker mode (`--texthooker`) and subtitle websocket behavior.
- Preserve mining/runtime options actions and trigger paths.

## Per-PR Required Automated Checks

- `pnpm run build`
- `pnpm run test:config`
- `pnpm run test:core`
- Current line gate script for milestone:
  - Example Gate 1: `pnpm run check:main-lines:gate1`

## Per-PR Manual Smoke Checks

- CLI:
  - `electron . --help` output is valid
  - `--start`, `--stop`, `--toggle` still route correctly
- Overlay:
  - visible overlay show/hide/toggle works
  - invisible overlay show/hide/toggle works
- Subtitle behavior:
  - subtitle updates still render
  - copy/mine shortcuts still function
- Integration:
  - runtime options palette opens
  - texthooker mode serves UI and can be opened

## Extraction Rules

- Move code verbatim first, refactor internals second.
- Keep temporary adapters/shims in `main.ts` until parity is verified.
- Limit each PR to one subsystem/risk area.
- If a regression appears, revert only that extraction slice and keep prior working structure.
