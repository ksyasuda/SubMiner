<!-- read_when: locating ownership for a runtime, feature, or integration -->

# Domain Ownership

Status: active
Last verified: 2026-03-13
Owner: Kyle Yasuda
Read when: you need to find the owner module for a behavior or test surface

## Runtime Domains

- Desktop app runtime: `src/main.ts`, `src/main/`, `src/core/services/`
- Overlay renderer: `src/renderer/`
- Launcher CLI: `launcher/`
- mpv plugin: `plugin/subminer/`

## Product / Integration Domains

- Config system: `src/config/`
- Overlay/window state: `src/core/services/overlay-*`, `src/main/overlay-*.ts`
- MPV runtime and protocol: `src/core/services/mpv*.ts`
- Subtitle/token pipeline: `src/core/services/tokenizer*`, `src/subtitle/`, `src/tokenizers/`
- Anki workflow: `src/anki-integration/`, `src/core/services/anki-jimaku*.ts`
- Immersion tracking: `src/core/services/immersion-tracker/`
- AniList tracking: `src/core/services/anilist/`, `src/main/runtime/composers/anilist-*`
- Jellyfin integration: `src/core/services/jellyfin*.ts`, `src/main/runtime/composers/jellyfin-*`
- Window trackers: `src/window-trackers/`
- Stats app: `stats/`
- Public docs site: `docs-site/`

## Ownership Heuristics

- Runtime wiring or dependency setup: start in `src/main/`
- Business logic or service behavior: start in `src/core/services/`
- UI interaction or overlay DOM behavior: start in `src/renderer/`
- Command parsing or mpv launch flow: start in `launcher/`
- User-facing docs: `docs-site/`
- Internal process/docs: `docs/`
