---
name: 'subminer-scrum-master'
description: 'Compatibility shim. Canonical SubMiner scrum-master workflow now lives in the repo-local subminer-workflow plugin.'
---

# Compatibility Shim

Canonical source:

- `plugins/subminer-workflow/skills/subminer-scrum-master/SKILL.md`

When this shim is invoked:

1. Read the canonical plugin-owned skill.
2. Follow the plugin-owned skill as the source of truth.
3. Do not duplicate workflow changes here; update the plugin-owned skill instead.

This shim exists so existing repo references and prompts keep resolving during the migration to the repo-local plugin workflow.
