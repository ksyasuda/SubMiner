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
bun run test:config       # Source-level config schema/validation tests
bun run test:launcher     # Launcher regression tests (config discovery + command routing)
bun run test:launcher:smoke:src # Launcher e2e smoke: launcher -> mpv IPC -> overlay start/stop wiring
bun run test:core         # Source-level core regression tests (default lane)
bun run test:subtitle     # Subtitle pipeline tests (build + run)
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
bun run test:subtitle:dist    # subtitle dist lane (currently placeholder)
```

## Maintainability Guardrails

Run guardrails locally before opening a PR:

```bash
bun run check:file-budgets:strict
bun run check:main-fanin:strict
bun run check:runtime-cycles:strict
```

Expected success output includes:

- `[OK] hotspot budget check (strict) — no hotspots over limit`
- `[OK] main runtime fan-in (strict) — ...`
- `[OK] runtime cycle check (strict) - ... no cycles detected`

Hotspot budget baselines (2026-02-22):

| File                                      | Baseline LOC | Limit LOC |
| ----------------------------------------- | ------------ | --------- |
| `src/main.ts`                             | 2978         | 3150      |
| `src/config/service.ts`                   | 116          | 140       |
| `src/core/services/tokenizer.ts`          | 232          | 260       |
| `launcher/main.ts`                        | 101          | 130       |
| `src/config/resolve.ts`                   | 33           | 55        |
| `src/main/runtime/composers/contracts.ts` | 13           | 30        |

Troubleshooting guardrail failures:

- File budget hotspot failure: split logic by runtime/domain boundaries and keep orchestration facades thin.
- Main fan-in failure: move runtime imports behind `src/main/runtime/domains/*` or composer barrels.
- Runtime cycle failure: break bidirectional imports by extracting shared helpers into leaf modules.
- Optional broad-budget audit (legacy/global): run `bun run scripts/check-file-budgets.ts --strict --enforce-global`.

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

- To add/change a config default, edit the matching domain file in `src/config/definitions/defaults-*.ts`.
- To add/change config option metadata, edit the matching domain file in `src/config/definitions/options-*.ts`.
- To add/change generated config template blocks/comments, update `src/config/definitions/template-sections.ts`.
- Keep `src/config/definitions.ts` as the composed public API (`DEFAULT_CONFIG`, registries, template export) that wires domain modules together.
- Overlay window/visibility state is owned by `src/core/services/overlay-manager.ts`.
- Runtime architecture/module-boundary conventions are documented in [Architecture](/architecture); keep contributor changes aligned with that canonical guide.
- Linux packaged desktop launches pass `--background` using electron-builder `build.linux.executableArgs` in `package.json`.
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
