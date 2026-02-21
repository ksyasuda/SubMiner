# Development

## Prerequisites

- [Node.js](https://nodejs.org/) (LTS)
- [Bun](https://bun.sh)

## Setup

```bash
git clone https://github.com/ksyasuda/SubMiner.git
cd SubMiner
make deps
# or manually:
bun install
cd vendor/texthooker-ui && pnpm install --frozen-lockfile
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
bash scripts/verify-generated-launcher.sh
```

## Running Locally

```bash
bun run dev    # builds + launches with --start --dev
electron . --start --dev --log-level debug   # equivalent Electron launch with verbose logging
electron . --background                       # tray/background mode, minimal default logging
```

## Testing

```bash
bun run test:config       # Config schema and validation tests (build + run)
bun run test:core         # Core service tests (~67 tests) (build + run)
bun run test:subtitle     # Subtitle pipeline tests (build + run)
```

All legacy test commands build first, then run via Node's built-in test runner (`node --test`).

For faster iteration while editing test code:

```bash
bun run build             # one-time compile
bun run test:config:dist  # no rebuild
bun run test:core:dist    # no rebuild
bun run test:subtitle:dist # no rebuild
bun run test:fast         # run all tests without rebuild (assumes build is already current)
```

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
| `make install`         | Install platform artifacts (wrapper, theme, AppImage/app bundle) |
| `make install-plugin`  | Install mpv Lua plugin and config                                |
| `make deps`            | Install JS dependencies (root + texthooker-ui)                   |
| `make generate-config` | Generate default config from centralized registry                |
| `make docs-dev`        | Run VitePress dev server                                         |

## Contributor Notes

- To add or change a config option, update `src/config/definitions.ts` first. Defaults, runtime-option metadata, and generated `config.example.jsonc` are derived from this centralized source.
- Overlay window/visibility state is owned by `src/core/services/overlay-manager.ts`.
- Main process composition is split across `src/main/` modules plus focused runtime composers under `src/main/runtime/composers/*` (for example AniList tracking and MPV runtime assembly clusters).
- Runtime domain imports for `src/main.ts` should route through `src/main/runtime/domains/*`; shared domain access point is `src/main/runtime/registry.ts`.
- Linux packaged desktop launches pass `--background` using electron-builder `build.linux.executableArgs` in `package.json`.
- MPV service has been split into transport, protocol, state, and properties layers in `src/core/services/`.
- Prefer direct inline deps objects in `src/main/` modules for simple pass-through wiring.
- Add a helper/adapter service only when it performs meaningful adaptation, validation, or reuse (not identity mapping).
- See [Architecture](/architecture) for the composition model and extension rules.

## Environment Variables

| Variable                          | Description                                                |
| --------------------------------- | ---------------------------------------------------------- |
| `SUBMINER_APPIMAGE_PATH`          | Override AppImage location for subminer script             |
| `SUBMINER_BINARY_PATH`            | Alias for `SUBMINER_APPIMAGE_PATH`                         |
| `SUBMINER_YT_SUBGEN_MODE`         | Override `youtubeSubgen.mode` for launcher                 |
| `SUBMINER_WHISPER_BIN`            | Override `youtubeSubgen.whisperBin` for launcher           |
| `SUBMINER_WHISPER_MODEL`          | Override `youtubeSubgen.whisperModel` for launcher         |
| `SUBMINER_YT_SUBGEN_OUT_DIR`      | Override generated subtitle output directory               |
| `SUBMINER_YT_SUBGEN_AUDIO_FORMAT` | Override extraction format used for whisper fallback       |
| `SUBMINER_YT_SUBGEN_KEEP_TEMP`    | Set to `1` to keep temporary subtitle-generation workspace |
