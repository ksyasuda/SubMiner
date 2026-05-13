---
name: 'subminer-change-verification'
description: 'Compatibility shim. Canonical SubMiner change verification workflow now lives in the repo-local subminer-workflow plugin.'
---

# Compatibility Shim

Canonical source:

- `plugins/subminer-workflow/skills/subminer-change-verification/SKILL.md`

Canonical helper scripts:

- `plugins/subminer-workflow/skills/subminer-change-verification/scripts/classify_subminer_diff.sh`
- `plugins/subminer-workflow/skills/subminer-change-verification/scripts/verify_subminer_change.sh`

When this shim is invoked:

1. Read the canonical plugin-owned skill.
2. Follow the plugin-owned skill as the source of truth.
3. Use the wrapper scripts in this shim directory only for compatibility with existing commands and docs.
4. Do not duplicate workflow changes here; update the plugin-owned skill and scripts instead.
