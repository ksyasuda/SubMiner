<!-- read_when: locating ownership for a runtime, feature, or integration -->

# Domain Ownership

Status: active
Last verified: 2026-07-15
Owner: Kyle Yasuda
Read when: you need to find the owner module for a behavior or test surface

## Runtime Domains

- Desktop app runtime: `src/main.ts`, `src/main/`, `src/core/services/`
- Overlay renderer: `src/renderer/`
- Launcher CLI: `launcher/`
- mpv plugin: `plugin/subminer/`

## Product / Integration Domains

- Config system: `src/config/`; Anki resolution is composed by
  `src/config/resolve/anki-connect.ts` from focused resolvers in
  `src/config/resolve/anki-connect/`
- Overlay/window state: `src/core/services/overlay-*`, `src/main/overlay-*.ts`
- MPV runtime and protocol: `src/core/services/mpv*.ts`
- Subtitle/token pipeline: `src/core/services/subtitle-*.ts`, `src/core/services/tokenizer*`, `src/core/services/tokenizer/`, `src/subsync/`
- Anki workflow: `src/anki-integration/`, `src/core/services/anki-jimaku*.ts`
- Immersion tracking: `src/core/services/immersion-tracker/`
  Includes stats storage/query schema such as `imm_videos`, `imm_media_art`, and `imm_youtube_videos` for per-video and YouTube-specific library metadata.
- AniList tracking + character dictionary: `src/core/services/anilist/`, `src/main/runtime/composers/anilist-*`, `src/main/character-dictionary-runtime.ts`, `src/main/character-dictionary-runtime/`
- Jellyfin integration: `src/core/services/jellyfin*.ts`, `src/main/runtime/composers/jellyfin-*`
- Window trackers: `src/window-trackers/`
- Stats HTTP app: `src/core/services/stats-server.ts`, with route groups and shared route support
  in `src/core/services/stats-server/`
- Stats SPA: `stats/`
- Public docs site: `docs-site/`

## Shared Contract Entry Points

- Config + app-state contracts: `src/types/config.ts`
- Subtitle/token/media annotation contracts: `src/types/subtitle.ts`
- Runtime/window/controller/Electron bridge contracts: `src/types/runtime.ts`
- Anki-specific contracts: `src/types/anki.ts`
- External integration contracts: `src/types/integrations.ts`
- Runtime-option contracts: `src/types/runtime-options.ts`
- Settings UI contracts: `src/types/settings.ts`
- Session-binding contracts: `src/types/session-bindings.ts`
- Stats HTTP wire contracts: `src/types/stats-wire.ts`, `src/types/stats-http-contract.ts`
- Compatibility-only barrel: `src/types.ts`

## Ownership Heuristics

- Runtime wiring or dependency setup: start in `src/main/`
- Business logic or service behavior: start in `src/core/services/`
- UI interaction or overlay DOM behavior: start in `src/renderer/`
- Command parsing or mpv launch flow: start in `launcher/`
- Shared contract changes: add or edit the narrowest `src/types/<domain>.ts` entrypoint; only touch `src/types.ts` for compatibility exports.
- User-facing docs: `docs-site/`
- Internal process/docs: `docs/`
