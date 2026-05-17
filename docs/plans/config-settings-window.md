# Config Settings Window

read_when: changing config UI, config save behavior, or config docs

## Intent

Add a dedicated Electron settings window for editing canonical config values without exposing the historical layout mistakes in `config.jsonc`.

The UI groups options by workflow:

- Viewing
- Mining & Anki
- Playback & Sources
- Input
- Integrations
- Tracking & App
- Advanced

Each field maps back to its current raw config path. The presentation layer must stay separate from generated config-template sections.

## Sources

- Canonical defaults: `DEFAULT_CONFIG`
- Existing option descriptions/enums: `CONFIG_OPTION_REGISTRY`
- UI registry: `src/config/settings/registry.ts`
- JSONC save path: `src/config/settings/jsonc-edit.ts`
- Window runtime: `src/main/runtime/config-settings-window.ts`

## Save Contract

Settings writes use `jsonc-parser.modify`, not `JSON.stringify`.

Required behavior:

- Preserve comments, trailing commas, unrelated keys, and hidden legacy keys.
- Reset removes the explicit path so defaults resolve normally.
- Validate the candidate config before writing.
- Reject warnings caused by modified fields.
- Preserve unrelated existing warnings and return them in the snapshot.
- Write atomically, reload `ConfigService`, classify with existing hot-reload logic, and apply live changes where supported.
- Never return secret values to the renderer; snapshots only expose configured/not-configured state.

## Hidden Compatibility Keys

Do not expose these as first-class controls:

- `ankiConnect.deck`
- Legacy top-level Anki migration fields such as `wordField`, `audioField`, media-generation aliases, and behavior aliases
- Legacy `ankiConnect.nPlusOne.*` aliases except canonical `nPlusOne.nPlusOne` and `nPlusOne.minSentenceWords`
- Deprecated Lapis sentence-card fields
- `youtubeSubgen.primarySubLanguages`
- `anilist.characterDictionary.refreshTtlHours`
- `anilist.characterDictionary.evictionPolicy`
- `jellyfin.accessToken`
- `jellyfin.userId`
- `controller.buttonIndices` as a normal editable field

## Verification

Minimum targeted checks:

- `bun test src/config/settings/registry.test.ts src/config/settings/jsonc-edit.test.ts src/settings/settings-model.test.ts src/main/runtime/config-settings-window.test.ts`
- `bun run test:config`
- `bun run typecheck`
- `bun run build`

If docs changed:

- `bun run docs:test`
- `bun run docs:build`
