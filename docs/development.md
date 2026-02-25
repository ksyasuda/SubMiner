# Development

## Prerequisites

- [Bun](https://bun.sh)

## Setup

```bash
git clone --recurse-submodules https://github.com/ksyasuda/SubMiner.git
cd SubMiner
# if you cloned without --recurse-submodules:
git submodule update --init --recursive

make deps
# or manually:
bun install
(cd vendor/texthooker-ui && bun install --frozen-lockfile)
```

## Building

```bash
# TypeScript compile (fast, for development)
bun run build

# Generate launcher wrapper artifact
make build-launcher
# output: dist/launcher/subminer

# Full platform build (includes texthooker-ui + AppImage/DMG)
make build

# Platform-specific builds
make build-linux          # Linux AppImage
make build-macos          # macOS DMG + ZIP (signed)
make build-macos-unsigned # macOS DMG + ZIP (unsigned)
```

## Launcher Artifact Workflow

- Source of truth: `launcher/*.ts`
- Generated output: `dist/launcher/subminer`
- Do not hand-edit generated launcher output.
- Repo-root `./subminer` is a stale artifact path and is rejected by verification checks.
- Install targets (`make install-linux`, `make install-macos`) copy from `dist/launcher/subminer`.

Verify the workflow:

```bash
make build-launcher
dist/launcher/subminer --help >/dev/null
bash scripts/verify-generated-launcher.sh
```

## Running Locally

```bash
bun run dev    # builds + launches with --start --dev
electron . --start --dev --log-level debug   # equivalent Electron launch with verbose logging
electron . --background                       # tray/background mode, minimal default logging
make dev-start                                # build + launch via Makefile
make dev-watch                                # watch TS + renderer and launch Electron (faster edit loop)
make dev-watch-macos                          # same as dev-watch, forcing --backend macos
```

For mpv-plugin-driven testing without exporting `SUBMINER_BINARY_PATH` each run, set a one-time
dev binary path in `~/.config/mpv/script-opts/subminer.conf`:

```ini
binary_path=/absolute/path/to/SubMiner/scripts/subminer-dev.sh
```

## Testing

CI-equivalent local gate:

```bash
bun run tsc --noEmit
bun run test:fast
bun run test:launcher:smoke:src
bun run build
bun run test:smoke:dist
bun run docs:build
```

Common focused commands:

```bash
bun run test:config       # Source-level config schema/validation tests
bun run test:launcher     # Launcher regression tests (config discovery + command routing)
bun run test:launcher:smoke:src # Launcher e2e smoke: launcher -> mpv IPC -> overlay start/stop wiring
bun run test:core         # Source-level core regression tests (default lane)
bun run test:fast         # Source-level config + core lane (no build prerequisite)
```

Dist-level tests are now an explicit smoke lane used to validate compiled/runtime assumptions.

Launcher smoke artifacts are written to `.tmp/launcher-smoke` locally and uploaded by CI/release workflows when the smoke step fails.

Smoke and optional deep dist commands:

```bash
bun run build                 # compile dist artifacts
bun run test:smoke:dist       # explicit smoke scope for compiled runtime
bun run test:config:dist      # optional full dist config suite
bun run test:core:dist        # optional full dist core suite
```

`bun run test:subtitle` and `bun run test:subtitle:dist` are currently placeholders and do not run an active suite.

## Config Generation

```bash
# Generate default config to ~/.config/SubMiner/config.jsonc
make generate-config

# Regenerate the repo's config.example.jsonc from centralized defaults
make generate-example-config
# or: bun run generate:config-example
```

## Documentation Site

The docs use [VitePress](https://vitepress.dev/):

```bash
make docs-dev     # Dev server at http://localhost:5173
make docs         # Build static output
make docs-preview # Preview built site at http://localhost:4173
```

## Makefile Reference

Run `make help` for a full list of targets. Key ones:

| Target                 | Description                                                      |
| ---------------------- | ---------------------------------------------------------------- |
| `make build`           | Build platform package for detected OS                           |
| `make build-launcher`  | Generate Bun launcher wrapper at `dist/launcher/subminer`        |
| `make install`         | Install platform artifacts (wrapper, theme, AppImage/app bundle) |
| `make install-plugin`  | Install mpv Lua plugin and config                                |
| `make deps`            | Install JS dependencies (root + texthooker-ui)                   |
| `make generate-config` | Generate default config from centralized registry                |
| `make docs-dev`        | Run VitePress dev server                                         |

## Contributor Notes

- To add/change a config default, edit the matching domain file in `src/config/definitions/defaults-*.ts`.
- To add/change config option metadata, edit the matching domain file in `src/config/definitions/options-*.ts`.
- To add/change generated config template blocks/comments, update `src/config/definitions/template-sections.ts`.
- Keep `src/config/definitions.ts` as the composed public API (`DEFAULT_CONFIG`, registries, template export) that wires domain modules together.
- Overlay window/visibility state is owned by `src/core/services/overlay-manager.ts`.
- Runtime architecture/module-boundary conventions are documented in [Architecture](/architecture); keep contributor changes aligned with that canonical guide.
- Linux packaged desktop launches pass `--background` using electron-builder `build.linux.executableArgs` in `package.json`.
- Prefer direct inline deps objects in `src/main/` modules for simple pass-through wiring.
- Add a helper/adapter service only when it performs meaningful adaptation, validation, or reuse (not identity mapping).

## Environment Variables

| Variable                           | Description                                                                    |
| ---------------------------------- | ------------------------------------------------------------------------------ |
| `SUBMINER_APPIMAGE_PATH`           | Override SubMiner app binary path for launcher playback commands               |
| `SUBMINER_BINARY_PATH`             | Alias for `SUBMINER_APPIMAGE_PATH`                                             |
| `SUBMINER_ROFI_THEME`              | Override rofi theme path for launcher picker                                   |
| `SUBMINER_LOG_LEVEL`               | Override app logger level (`debug`, `info`, `warn`, `error`)                   |
| `SUBMINER_MPV_LOG`                 | Override mpv/app shared log file path                                          |
| `SUBMINER_YT_SUBGEN_MODE`          | Override `youtubeSubgen.mode` for launcher                                     |
| `SUBMINER_WHISPER_BIN`             | Override `youtubeSubgen.whisperBin` for launcher                               |
| `SUBMINER_WHISPER_MODEL`           | Override `youtubeSubgen.whisperModel` for launcher                             |
| `SUBMINER_YT_SUBGEN_OUT_DIR`       | Override generated subtitle output directory                                   |
| `SUBMINER_YT_SUBGEN_AUDIO_FORMAT`  | Override extraction format used for whisper fallback                           |
| `SUBMINER_YT_SUBGEN_KEEP_TEMP`     | Set to `1` to keep temporary subtitle-generation workspace                     |
| `SUBMINER_JIMAKU_API_KEY`          | Override Jimaku API key for launcher subtitle downloads                        |
| `SUBMINER_JIMAKU_API_KEY_COMMAND`  | Command used to resolve Jimaku API key at runtime                              |
| `SUBMINER_JIMAKU_API_BASE_URL`     | Override Jimaku API base URL                                                   |
| `SUBMINER_JELLYFIN_ACCESS_TOKEN`   | Override Jellyfin access token (used before stored encrypted session fallback) |
| `SUBMINER_JELLYFIN_USER_ID`        | Optional Jellyfin user ID override                                             |
| `SUBMINER_SKIP_MACOS_HELPER_BUILD` | Set to `1` to skip building the macOS helper binary during `bun run build`     |
