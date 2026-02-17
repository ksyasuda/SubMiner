# Development

## Prerequisites

- [Node.js](https://nodejs.org/) (LTS)
- [pnpm](https://pnpm.io/)
- [Bun](https://bun.sh) (for the `subminer` wrapper script)

## Setup

```bash
git clone https://github.com/ksyasuda/SubMiner.git
cd SubMiner
make deps
# or manually:
pnpm install
pnpm -C vendor/texthooker-ui install
```

## Building

```bash
# TypeScript compile (fast, for development)
pnpm run build

# Full platform build (includes texthooker-ui + AppImage/DMG)
make build

# Platform-specific builds
make build-linux          # Linux AppImage
make build-macos          # macOS DMG + ZIP (signed)
make build-macos-unsigned # macOS DMG + ZIP (unsigned)
```

## Running Locally

```bash
pnpm run dev    # builds + launches with --start --dev
electron . --start --dev --log-level debug   # equivalent Electron launch with verbose logging
```

## Testing

```bash
pnpm run test:config       # Config schema and validation tests (build + run)
pnpm run test:core         # Core service tests (~67 tests) (build + run)
pnpm run test:subtitle     # Subtitle pipeline tests (build + run)
```

All legacy test commands build first, then run via Node's built-in test runner (`node --test`).

For faster iteration while editing test code:

```bash
pnpm run build             # one-time compile
pnpm run test:config:dist  # no rebuild
pnpm run test:core:dist    # no rebuild
pnpm run test:subtitle:dist # no rebuild
pnpm run test:fast         # run all tests without rebuild (assumes build is already current)
```

## Config Generation

```bash
# Generate default config to ~/.config/SubMiner/config.jsonc
make generate-config

# Regenerate the repo's config.example.jsonc from centralized defaults
make generate-example-config
# or: pnpm run generate:config-example
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

| Target | Description |
| --- | --- |
| `make build` | Build platform package for detected OS |
| `make install` | Install platform artifacts (wrapper, theme, AppImage/app bundle) |
| `make install-plugin` | Install mpv Lua plugin and config |
| `make deps` | Install JS dependencies (root + texthooker-ui) |
| `make generate-config` | Generate default config from centralized registry |
| `make docs-dev` | Run VitePress dev server |

## Contributor Notes

- To add or change a config option, update `src/config/definitions.ts` first. Defaults, runtime-option metadata, and generated `config.example.jsonc` are derived from this centralized source.
- Overlay window/visibility state is owned by `src/core/services/overlay-manager.ts`.
- Main process composition is now split across `src/main/` modules (`startup.ts`, `app-lifecycle.ts`, `startup-lifecycle.ts`, `state.ts`, `ipc-runtime.ts`, `cli-runtime.ts`, `overlay-runtime.ts`, `subsync-runtime.ts`).
- MPV service has been split into transport, protocol, state, and properties layers in `src/core/services/`.
- Prefer direct inline deps objects in `src/main/` modules for simple pass-through wiring.
- Add a helper/adapter service only when it performs meaningful adaptation, validation, or reuse (not identity mapping).
- See [Architecture](/architecture) for the composition model and extension rules.

## Environment Variables

| Variable | Description |
| --- | --- |
| `SUBMINER_APPIMAGE_PATH` | Override AppImage location for subminer script |
| `SUBMINER_BINARY_PATH` | Alias for `SUBMINER_APPIMAGE_PATH` |
| `SUBMINER_YT_SUBGEN_MODE` | Override `youtubeSubgen.mode` for launcher |
| `SUBMINER_WHISPER_BIN` | Override `youtubeSubgen.whisperBin` for launcher |
| `SUBMINER_WHISPER_MODEL` | Override `youtubeSubgen.whisperModel` for launcher |
| `SUBMINER_YT_SUBGEN_OUT_DIR` | Override generated subtitle output directory |
| `SUBMINER_YT_SUBGEN_AUDIO_FORMAT` | Override extraction format used for whisper fallback |
| `SUBMINER_YT_SUBGEN_KEEP_TEMP` | Set to `1` to keep temporary subtitle-generation workspace |
