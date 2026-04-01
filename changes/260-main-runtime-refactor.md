type: internal
area: main

- Split `src/main.ts` into domain runtime wrappers and startup sequencing helpers.
- Removed the last direct `src/main/runtime/*-main-deps.ts` imports from `src/main.ts`.
- Kept startup behavior and IPC contracts stable while reducing composition-root size.
