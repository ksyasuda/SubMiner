# Internal Knowledge Base Restructure Design

**Problem:** `AGENTS.md` currently carries too much project detail while deeper internal guidance is either missing from `docs/` or mixed into `docs-site/`, which should stay user-facing. Agents and contributors need a stable entrypoint plus an internal system of record with progressive disclosure and mechanical enforcement.

**Goals:**
- Make `AGENTS.md` a short table of contents, not an encyclopedia.
- Establish `docs/` as the internal system of record for architecture, workflow, and knowledge-base conventions.
- Keep `docs-site/` user-facing only.
- Add lightweight enforcement so the split is maintained mechanically.
- Preserve existing build/test/release guidance while moving canonical internal pointers into `docs/`.

**Non-Goals:**
- Rework product/user docs information architecture beyond boundary cleanup.
- Build a custom documentation generator.
- Solve plan lifecycle cleanup in this change.
- Reorganize unrelated runtime code or existing feature docs.

## Recommended Approach

Create a small internal knowledge-base structure under `docs/`, rewrite `AGENTS.md` into a compact map that points into that structure, and add a repo-level test that enforces required internal docs and boundary rules. Keep `docs-site/` focused on public/product documentation and strip out any claims that it is the canonical source for internal architecture or workflow.

## Information Architecture

Proposed internal layout:

```text
docs/
  README.md
  architecture/
    README.md
    domains.md
    layering.md
  knowledge-base/
    README.md
    core-beliefs.md
    catalog.md
    quality.md
  workflow/
    README.md
    planning.md
    verification.md
  plans/
    ...
```

Key rules:
- `AGENTS.md` links to `docs/README.md` plus a small set of core entrypoints.
- `docs/README.md` acts as the internal KB home page.
- Internal docs include lightweight metadata via explicit section fields:
  - `Status`
  - `Last verified`
  - `Owner`
  - `Read when`
- `docs-site/` remains public/user-facing and may link to internal docs only as contributor references, not as canonical internal source-of-truth pages.

## Enforcement

Add a repo-level knowledge-base test that validates:

- `AGENTS.md` links to required internal KB docs.
- `AGENTS.md` stays below a capped line count.
- Required internal docs exist.
- Internal KB docs include the expected metadata fields.
- `docs-site/development.md` and `docs-site/architecture.md` point internal readers to `docs/`.
- `AGENTS.md` does not treat `docs-site/` pages as the canonical internal source of truth.

Keep the new test outside `docs-site/` so internal/public boundaries stay clear.

## Migration Plan

1. Write the design and implementation plan docs.
2. Rewrite `AGENTS.md` as a compact map.
3. Create the internal KB entrypoints under `docs/`.
4. Update `docs-site/` contributor docs to reference `docs/` for internal guidance.
5. Add a repo-level test plus package/CI wiring.
6. Run docs and targeted repo verification.

## Verification Strategy

Primary commands:
- `bun run docs:test`
- `bun run test:fast`
- `bun run docs:build`

Focused review:
- read `AGENTS.md` start-to-finish and confirm it behaves like a table of contents
- inspect the new `docs/README.md` navigation and cross-links
- confirm `docs-site/` still reads as user-facing documentation
