# IPC + Runtime Contracts

## Core Surfaces

- `src/shared/ipc/contracts.ts`: canonical channel names and payload contracts
- `src/shared/ipc/validators.ts`: runtime payload validation/parsing
- `src/preload.ts`: renderer bridge to approved IPC endpoints
- `src/main/ipc-runtime.ts`: main-process registration and handler routing
- `src/core/services/ipc.ts`: service-level invoke handling and guardrails
- `src/core/services/anki-jimaku-ipc.ts`: integration-specific IPC boundary
- `src/main/cli-runtime.ts`: CLI/runtime command boundary

## Contract Rules

- Use shared contract constants; avoid ad-hoc literal channel strings.
- Validate `invoke` payloads before domain/service logic.
- Return structured failures (`{ ok: false, error }`) where possible.
- Keep payloads narrow and explicit.
- Update contracts, validators, preload types, and handlers in the same change when shape changes.

## Add a New IPC Action

1. Add channel in `src/shared/ipc/contracts.ts`.
2. Add or extend validator in `src/shared/ipc/validators.ts`.
3. Expose typed bridge method in `src/preload.ts`.
4. Register handler in `src/main/ipc-runtime.ts` (or relevant domain runtime module).
5. Add tests for valid and malformed payload cases in `src/core/services/*`.
6. Update renderer tests when behavior or state transitions change.

## Runtime State Notes

- Prefer runtime/domain composition via `src/main/runtime/composers/*` and `src/main/runtime/domains/*`.
- Route shared mutable state updates through transition helpers in `src/main/state.ts` for migrated domains.
- Keep IPC handlers thin; avoid embedding business logic in transport wiring.

## Troubleshooting

- Unknown payload in handler: confirm validator is applied before handler/service call.
- Renderer invoke fails: verify preload bridge method and channel registration.
- Contract drift: compare shared contract, validator, preload bridge, and main handler signatures together.

## Related Docs

- [Architecture](/architecture)
- [Development](/development)
- [Configuration](/configuration)
- [Troubleshooting](/troubleshooting)
