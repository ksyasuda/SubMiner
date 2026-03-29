<!-- read_when: starting implementation, deciding whether to plan, or checking handoff expectations -->

# Workflow

Status: active  
Last verified: 2026-03-13  
Owner: Kyle Yasuda  
Read when: planning or executing nontrivial work in this repo

This section is the internal workflow map for contributors and agents.

## Read Next

- [Planning](./planning.md) - when to write a lightweight plan vs a full execution plan
- [Verification](./verification.md) - maintained test/build lanes and handoff gate
- [Agent Plugins](./agent-plugins.md) - repo-local plugin ownership for agent workflow skills
- [Release Guide](../RELEASING.md) - tagged release workflow

## Default Flow

1. Read the smallest relevant docs from `docs/`.
2. Decide whether the work needs a written plan.
3. Implement in small, reviewable edits.
4. Run the cheapest sufficient verification lane.
5. Escalate to the full maintained gate before handoff when the change is substantial.

## Boundaries

- Internal process lives in `docs/`.
- Public/product docs live in `docs-site/`.
- Generated artifacts are never edited by hand.
