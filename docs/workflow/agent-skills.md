<!-- read_when: using or modifying repo-local agent skills -->

# Agent Skills

Status: active
Last verified: 2026-08-13
Owner: Kyle Yasuda
Read when: using, adding, or changing a repo-local agent workflow skill

## Canonical Skills

- `.agents/skills/subminer-change-verification/`
  - Selects the cheapest sufficient repo-native verification lane.
  - Defers command ownership to `package.json` and `docs/workflow/verification.md`.

Repo-local workflows stay as standalone skills. Do not add plugin packaging, marketplace metadata, or compatibility shims unless the workflow is intentionally being distributed beyond this repository.

## Rules

- Keep each skill focused on one repeatable repository task.
- Prefer instructions over helper scripts unless deterministic tooling provides clear value beyond existing package commands.
- Keep trigger descriptions narrow enough to avoid invoking skills for unrelated requests.
- Update this page and the documentation catalog when skill ownership changes.

## Verification

For skill or internal workflow documentation changes, run:

```bash
bun run test:docs:kb
```
