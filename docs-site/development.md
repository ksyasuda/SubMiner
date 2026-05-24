# Building & Testing

For internal architecture/workflow guidance, use `docs/README.md` at the repo root. This page stays focused on contributor-facing build and test commands.

## Prerequisites

- [Bun](https://bun.sh)

## Setup

```bash
git clone --recurse-submodules https://github.com/ksyasuda/SubMiner.git
cd SubMiner
# if you cloned without --recurse-submodules:
git submodule update --init --recursive

bun install
(cd stats && bun install --frozen-lockfile)
(cd vendor/texthooker-ui && bun install --frozen-lockfile)
```

`make deps` is still available as a convenience wrapper around the same dependency install flow.

## Building

```bash
# Main app build
bun run build

# Platform packages
bun run build:appimage      # Linux AppImage
bun run build:mac           # macOS DMG + ZIP (signed)
bun run build:mac:unsigned  # macOS DMG + ZIP (unsigned)
bun run build:win           # Windows NSIS installer + ZIP

# Optional launcher artifact only
make build-launcher
# output: dist/launcher/subminer
```

`bun run build` includes the Yomitan build step. It builds the bundled Chrome extension directly from the `vendor/subminer-yomitan` submodule into `build/yomitan` using Bun.

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
dev binary path with `mpv.subminerBinaryPath` in your SubMiner config. The launcher injects it into
the mpv plugin at runtime:

```json
{
  "mpv": {
    "subminerBinaryPath": "/absolute/path/to/SubMiner/scripts/subminer-dev.sh"
  }
}
```

## Testing

Default lanes:

```bash
bun run test           # alias for test:fast
bun run test:fast      # default fast lane
bun run test:full      # maintained source + launcher-unit + runtime compat surface
bun run test:runtime:compat # compiled/runtime compatibility slice only
bun run test:env       # launcher/plugin + env-sensitive verification
bun run test:immersion:sqlite # SQLite persistence lane
bun run test:subtitle  # maintained alass/ffsubsync subtitle surface
```

- `bun run test` and `bun run test:fast` cover config/core suites plus representative entry/runtime, Anki integration, release-workflow coverage, typecheck, and runtime-registry checks.
- `bun run test:full` is the maintained full surface: Bun-compatible `src/**` discovery, Bun-compatible launcher unit discovery, and the compiled/runtime compatibility lane for suites routed through `dist/**`.
- `bun run test:runtime:compat` covers the compiled/runtime slice directly: `ipc`, `anki-jimaku-ipc`, `overlay-manager`, `config-validation`, `startup-config`, and `registry`.
- `bun run test:env` covers environment-sensitive checks: launcher smoke/plugin verification plus the Bun source SQLite lane.
- `bun run test:immersion:sqlite` is the reproducible persistence lane when you need real DB-backed SQLite coverage under Bun.

The Bun-managed discovery lanes intentionally exclude a small compiled/runtime-focused set: `src/core/services/ipc.test.ts`, `src/core/services/anki-jimaku-ipc.test.ts`, `src/core/services/overlay-manager.test.ts`, `src/main/config-validation.test.ts`, `src/main/runtime/startup-config.test.ts`, and `src/main/runtime/registry.test.ts`. `bun run test:runtime:compat` keeps them in the standard workflow via `dist/**`.

Suggested local gate before handoff:

```bash
bun run typecheck
bun run test:fast
bun run test:env
bun run build
bun run test:smoke:dist
```

If you changed docs in `docs-site/`, also run:

```bash
bun run docs:test
bun run docs:build
```

For production docs routing, run the versioned build:

```bash
bun run docs:build:versioned
```

The versioned build writes `.tmp/docs-versioned-site` with latest stable docs at `/`, development docs at `/main/`, and stable archives under `/v/<version>/`. Prerelease tags are skipped. Public assets from `docs-site/public/assets` are shared from root `/assets/` so large demo media is not duplicated into every version archive; generated VitePress CSS and JS assets stay under each version route. Stale `.tmp/docs-versioned-archive-cache` generations are pruned after a successful build, and intermediate `.tmp/docs-versioned-build` workspaces are removed.

Focused commands:

```bash
bun run test:config       # Source-level config schema/validation tests
bun run test:launcher     # Launcher regression tests (config discovery + command routing)
bun run test:core         # Source-level core regression tests (default lane)
bun run test:launcher:smoke:src # Launcher e2e smoke: launcher -> mpv IPC -> overlay start/stop wiring
bun run test:launcher:env:src # Launcher smoke + Lua plugin gate
bun run test:src          # Bun-managed maintained src/** discovery lane
bun run test:launcher:unit:src # Bun-managed maintained launcher unit lane
bun run test:immersion:sqlite:src # Bun source lane
```

Dist-level tests are now an explicit smoke lane used to validate compiled/runtime assumptions.

Launcher smoke artifacts are written to `.tmp/launcher-smoke` locally and uploaded by CI/release workflows when the smoke step fails.

Smoke and optional deep dist commands:

```bash
bun run build                 # compile dist artifacts
bun run test:immersion:sqlite # compile + run SQLite-backed immersion tests under Bun
bun run test:smoke:dist       # explicit smoke scope for compiled runtime
bun run test:config:dist      # optional full dist config suite
bun run test:core:dist        # optional full dist core suite
```

Use `bun run test:immersion:sqlite` when you need real DB-backed coverage for the immersion tracker.

## Formatting

Use the scoped formatter for normal app-repo work:

```bash
make pretty
bun run format:check:src
```

- `make pretty` runs the maintained Prettier allowlist only (`format:src`).
- `bun run format:check:src` checks the same scoped set without writing changes.
- `bun run format` remains the broad repo-wide Prettier command; use it intentionally.

## Config Generation

```bash
# Generate default config to ~/.config/SubMiner/config.jsonc (or %APPDATA%\SubMiner\config.jsonc on Windows)
bun run electron . --generate-config

# Regenerate the repo's config.example.jsonc from centralized defaults
bun run generate:config-example
```

Convenience wrappers still exist:

- `make generate-config`
- `make generate-example-config`

## Documentation Site

The docs site now lives in `docs-site/` inside the main repo.

From the SubMiner app repo:

```bash
bun --cwd docs-site install
bun run docs:dev     # Dev server at http://localhost:5173
bun run docs:build   # Production build into docs-site/.vitepress/dist
bun run docs:preview # Preview built site at http://localhost:4173
bun run docs:test    # Docs regression tests
```

Cloudflare Pages deploy settings:

- Git repo: `ksyasuda/SubMiner`
- Root directory: `docs-site`
- Build command: `bun run docs:build`
- Build output directory: `.vitepress/dist`
- Build watch paths: `docs-site/*`

Use Cloudflare's single `*` wildcard syntax for watch paths. `docs-site/*` covers nested docs-site changes in the repo; `docs-site/**` is not the correct Pages pattern and may skip docs-only pushes.

## Makefile Reference

Run `make help` for a full list of targets. Key ones:

| Target                      | Description                                                       |
| --------------------------- | ----------------------------------------------------------------- |
| `make build`                | Build platform package for detected OS                            |
| `make build-launcher`       | Generate Bun launcher wrapper at `dist/launcher/subminer`         |
| `make install`              | Install platform artifacts (wrapper, theme, AppImage/app bundle)  |
| `make deps`                 | Install JS dependencies (root + stats + texthooker-ui)            |
| `make pretty`               | Run scoped Prettier formatting for maintained source/config files |
| `make generate-config`      | Generate default config from centralized registry                 |
| `make build-linux`          | Convenience wrapper for Linux packaging                           |
| `make build-macos`          | Convenience wrapper for signed macOS packaging                    |
| `make build-macos-unsigned` | Convenience wrapper for unsigned macOS packaging                  |

## Contributor Notes

- To add/change a config default, edit the matching domain file in `src/config/definitions/defaults-*.ts`.
- To add/change config option metadata, edit the matching domain file in `src/config/definitions/options-*.ts`.
- To add/change generated config template blocks/comments, update `src/config/definitions/template-sections.ts`.
- Keep `src/config/definitions.ts` as the composed public API (`DEFAULT_CONFIG`, registries, template export) that wires domain modules together.
- Overlay window/visibility state is owned by `src/core/services/overlay-manager.ts`.
- Runtime architecture/module-boundary conventions are summarized in [Architecture](/architecture), with canonical internal guidance in `docs/architecture/README.md` at the repo root.
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
| `SUBMINER_JIMAKU_API_KEY`          | Override Jimaku API key for launcher subtitle downloads                        |
| `SUBMINER_JIMAKU_API_KEY_COMMAND`  | Command used to resolve Jimaku API key at runtime                              |
| `SUBMINER_JIMAKU_API_BASE_URL`     | Override Jimaku API base URL                                                   |
| `SUBMINER_JELLYFIN_ACCESS_TOKEN`   | Override Jellyfin access token (used before stored encrypted session fallback) |
| `SUBMINER_JELLYFIN_USER_ID`        | Optional Jellyfin user ID override                                             |
| `SUBMINER_SKIP_MACOS_HELPER_BUILD` | Set to `1` to skip building the macOS helper binary during `bun run build`     |
