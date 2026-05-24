<!-- read_when: starting substantial work in this repo or looking for the internal source of truth -->

# SubMiner Internal Docs

Status: active  
Last verified: 2026-05-23  
Owner: Kyle Yasuda  
Read when: you need internal architecture, workflow, verification, or release guidance

`docs/` is the internal system of record for agent and contributor knowledge. Start here, then drill into the smallest doc that fits the task.

## Start Here

- [Architecture](./architecture/README.md) - runtime map, domains, layering rules
- [Workflow](./workflow/README.md) - planning, execution, verification expectations
- [Knowledge Base](./knowledge-base/README.md) - how docs are structured, maintained, and audited
- [Release Guide](./RELEASING.md) - tagged release checklist

## Fast Paths

- New feature or refactor: [Workflow](./workflow/README.md), then [Architecture](./architecture/README.md)
- Test/build/release work: [Verification](./workflow/verification.md), then [Release Guide](./RELEASING.md)
- Coverage lane selection or LCOV artifact path: [Verification](./workflow/verification.md)
- “What owns this behavior?”: [Domains](./architecture/domains.md)
- “Can these modules depend on each other?”: [Layering](./architecture/layering.md)
- “What doc should exist for this?”: [Catalog](./knowledge-base/catalog.md)

## Rules

- Treat `docs/` as canonical for internal guidance.
- Treat `docs-site/` as user-facing/public docs.
- Keep `AGENTS.md` short; deep detail belongs here.
- Update docs when behavior, architecture, or workflow meaningfully changes.
