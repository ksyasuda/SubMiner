<!-- read_when: adding dependencies, moving files, or reviewing architecture drift -->

# Layering Rules

Status: active  
Last verified: 2026-03-13  
Owner: Kyle Yasuda  
Read when: deciding whether a dependency direction is acceptable

## Preferred Dependency Flow

1. `src/main.ts`
2. `src/main/` composition and runtime adapters
3. `src/core/services/` focused services
4. `src/core/utils/` and other pure helpers

Renderer, launcher, plugin, and stats each keep their own local layering and should not become a grab bag for unrelated cross-runtime behavior.

## Rules

- Keep `src/main.ts` thin; wire, do not implement.
- Prefer injecting dependencies from `src/main/` instead of reaching outward from core services.
- Keep side effects explicit and close to composition boundaries.
- Put reusable business logic in focused services, not in top-level lifecycle files.
- Keep renderer concerns in `src/renderer/`; avoid leaking DOM behavior into main-process code.
- Treat `launcher/*.ts` as source of truth for the launcher. Never hand-edit `dist/launcher/subminer`.

## Smells

- `main.ts` grows because logic was not extracted
- service reaches directly into unrelated runtime state
- renderer code depends on main-process internals
- docs-site page becomes the only place internal architecture is explained
