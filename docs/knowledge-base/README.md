<!-- read_when: changing internal docs structure, adding guidance, or debugging doc drift -->

# Knowledge Base Rules

Status: active  
Last verified: 2026-03-13  
Owner: Kyle Yasuda  
Read when: maintaining the internal doc system itself

This section defines how the internal knowledge base is organized and maintained.

## Read Next

- [Core Beliefs](./core-beliefs.md) - agent-first operating principles
- [Catalog](./catalog.md) - indexed docs and verification status
- [Quality](./quality.md) - current doc and architecture quality grades

## Policy

- `AGENTS.md` is an entrypoint only.
- `docs/` is the internal system of record.
- `docs-site/` is user-facing; do not treat it as canonical internal design or workflow storage.
- Internal docs should be short, cross-linked, and specific.
- Every core internal doc should include:
  - `Status`
  - `Last verified`
  - `Owner`
  - `Read when`

## Maintenance

- Update the relevant internal doc when behavior or workflow changes.
- Add new docs to the [Catalog](./catalog.md).
- Record architectural quality drift in [Quality](./quality.md).
- Keep stale docs obvious; do not leave ambiguity about whether a page is trustworthy.
