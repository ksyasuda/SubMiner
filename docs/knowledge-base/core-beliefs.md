<!-- read_when: deciding how much context to inject, where docs should live, or how agents should navigate the repo -->

# Core Beliefs

Status: active  
Last verified: 2026-03-13  
Owner: Kyle Yasuda  
Read when: making decisions about agent ergonomics, doc structure, or repository guidance

## Agent-First Principles

- Progressive disclosure beats giant injected context.
- `AGENTS.md` should map the territory, not duplicate it.
- Canonical internal guidance belongs in versioned docs near the code.
- Plans are first-class while active work is happening.
- Mechanical checks beat social convention when the boundary matters.
- Small focused docs are easier to trust, update, and verify.
- User-facing docs and internal operating docs should not blur together.

## What This Means Here

- Start from `AGENTS.md`, then move into `docs/`.
- Prefer links to canonical docs over repeating long instructions.
- Keep architecture and workflow docs in separate pages so updates stay targeted.
- When a page becomes long or multi-purpose, split it.
