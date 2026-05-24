<!-- read_when: assessing architecture health, doc gaps, or where cleanup effort should go next -->

# Quality Scorecard

Status: active  
Last verified: 2026-03-13  
Owner: Kyle Yasuda  
Read when: triaging internal quality gaps or deciding where follow-up work is needed

Grades are directional, not ceremonial. The point is to keep gaps visible.

## Product / Runtime Domains

| Area | Grade | Notes |
| --- | --- | --- |
| Desktop runtime composition | B | strong modularization; still easy for `main` wiring drift to reappear |
| Launcher CLI | B | focused surface; generated/stale artifact hazards need constant guarding |
| mpv plugin | B | modular, but Lua/runtime coupling still specialized |
| Overlay renderer | B | improved modularity; interaction complexity remains |
| Config system | A- | clear defaults/definitions split and good validation surface |
| Immersion / AniList / Jellyfin surfaces | B- | growing product scope; ownership spans multiple services |
| Internal docs system | B | new structure in place; needs habitual maintenance |
| Public docs site | B | strong user docs; must stay separate from internal KB |

## Architectural Layers

| Layer | Grade | Notes |
| --- | --- | --- |
| `src/main.ts` composition root | B | direction good; still needs vigilance against logic creep |
| `src/main/` runtime adapters | B | mostly clear; can accumulate wiring debt |
| `src/core/services/` | B+ | good extraction pattern; some domains remain broad |
| `src/renderer/` | B | cleaner than before; UI/runtime behavior still dense |
| `launcher/` | B | clear command boundaries |
| `docs/` internal KB | B | structure exists; enforcement now guards core rules |

## Current Gaps

- Some deep architecture detail still lives in `docs-site/architecture.md` and may merit later migration.
- Quality grading is manual and should be refreshed when major refactors land.
