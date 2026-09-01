<!-- read_when: you need to know what internal docs exist, whether they are current, or what should be updated -->

# Documentation Catalog

Status: active  
Last verified: 2026-08-13
Owner: Kyle Yasuda  
Read when: finding internal docs or checking verification status

| Area | Path | Status | Last verified | Notes |
| --- | --- | --- | --- | --- |
| KB home | `docs/README.md` | active | 2026-05-23 | internal entrypoint |
| Architecture index | `docs/architecture/README.md` | active | 2026-05-23 | top-level runtime map |
| Domain ownership | `docs/architecture/domains.md` | active | 2026-05-23 | runtime and feature ownership |
| Layering rules | `docs/architecture/layering.md` | active | 2026-05-23 | dependency direction and smells |
| Subtitle overlay priming | `docs/architecture/subtitle-overlay-priming.md` | active | 2026-06-01 | visible-overlay subtitle startup flow |
| KB rules | `docs/knowledge-base/README.md` | active | 2026-05-23 | maintenance policy |
| Core beliefs | `docs/knowledge-base/core-beliefs.md` | active | 2026-03-13 | agent-first principles |
| Quality scorecard | `docs/knowledge-base/quality.md` | active | 2026-03-13 | quality grades and gaps |
| Workflow index | `docs/workflow/README.md` | active | 2026-08-13 | execution map |
| Planning guide | `docs/workflow/planning.md` | active | 2026-05-23 | lightweight vs execution plans |
| Agent skills | `docs/workflow/agent-skills.md` | active | 2026-08-23 | repo-local workflow skill ownership |
| Verification guide | `docs/workflow/verification.md` | active | 2026-08-13 | maintained verification lanes |
| Release guide | `docs/RELEASING.md` | active | 2026-05-23 | release checklist |

## Update Rules

- Add a row when introducing a new core internal doc.
- Update `Status` and `Last verified` when a page is materially revised.
- If a page is known inaccurate, mark it stale immediately instead of leaving silent drift.
