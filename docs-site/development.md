# Building and testing

Build, run, and test SubMiner from source. Architecture and workflow rules live in the repo's internal docs, starting at [`docs/README.md`](https://github.com/ksyasuda/SubMiner/blob/main/docs/README.md). The lane-by-lane test guide is [`docs/workflow/verification.md`](https://github.com/ksyasuda/SubMiner/blob/main/docs/workflow/verification.md).

## Prerequisites

- [Bun](https://bun.sh), at the version pinned in `package.json`
- A system `lua` interpreter for the mpv plugin tests (`bun run test:launcher`, `bun run test:env`)
- macOS only: `bun run build` compiles a Swift window helper. Set `SUBMINER_SKIP_MACOS_HELPER_BUILD=1` to skip it.

## Setup

```bash
git clone --recurse-submodules https://github.com/ksyasuda/SubMiner.git
cd SubMiner
make deps
```

`make deps` initializes submodules and installs dependencies for the root, `stats/`, and `vendor/texthooker-ui`. The Yomitan submodule installs its own dependencies during `bun run build`.

## Build

```bash
bun run build               # app build, including bundled Yomitan from vendor/subminer-yomitan
bun run build:appimage      # Linux AppImage
bun run build:mac           # macOS DMG + ZIP (signed)
bun run build:mac:unsigned  # macOS DMG + ZIP (unsigned)
bun run build:win           # Windows NSIS installer + ZIP
make build-launcher         # launcher only, output: dist/launcher/subminer
```

The launcher source is `launcher/*.ts`. `dist/launcher/subminer` is generated, so never edit it by hand. The repo-root `./subminer` is a stale path and verification rejects it. `make install-linux` and `make install-macos` copy from `dist/launcher/subminer`. To check the launcher build:

```bash
make build-launcher
dist/launcher/subminer --help >/dev/null
bash scripts/verify-generated-launcher.sh
```

## Run locally

```bash
bun run dev                                  # build, then launch with --start --dev
make dev-watch                               # watch TS + renderer and relaunch Electron
make dev-watch-macos                         # same, forcing --backend macos
electron . --start --dev --log-level debug   # verbose launch of an existing build
electron . --background                      # tray/background mode
```

To test through the mpv plugin without exporting `SUBMINER_BINARY_PATH` each time, point `mpv.subminerBinaryPath` in your config at the dev script. The launcher passes it to the plugin at runtime:

```json
{
  "mpv": {
    "subminerBinaryPath": "/absolute/path/to/SubMiner/scripts/subminer-dev.sh"
  }
}
```

## Test

Run the handoff gate before submitting substantial changes:

```bash
bun run typecheck
bun run test:fast
bun run test:env
bun run build
bun run test:smoke:dist
```

For smaller changes, start with the cheapest lane that covers what you touched:

| Command                         | Covers                                                                   |
| ------------------------------- | ------------------------------------------------------------------------ |
| `bun run test` / `test:fast`    | All `src/**` tests, launcher unit tests, and `scripts/**` tests          |
| `bun run test:config`           | Config schema, defaults, and `config.example.jsonc` generation           |
| `bun run test:launcher`         | Launcher tests plus the Lua plugin tests                                 |
| `bun run test:env`              | Launcher e2e smoke, Lua plugin tests, SQLite immersion tests from source |
| `bun run test:scripts`          | Build and release scripts under `scripts/**`                             |
| `bun run test:stats`            | Stats dashboard UI under `stats/src/**`                                  |
| `bun run test:runtime:compat`   | Compiled-runtime smoke against `dist/` (run `bun run build` first)       |
| `bun run test:immersion:sqlite` | Compiles, then runs the SQLite-backed immersion tracker tests            |
| `bun run test:subtitle`         | alass/ffsubsync subtitle sync                                            |
| `bun run test:docs:kb`          | Internal docs, `AGENTS.md`, and repo skills                              |

Lane membership is defined in `scripts/test-lanes.ts` and discovered by directory, so a new test file joins its lane automatically. Do not hand-list test files in `package.json`. `scripts/run-test-lane.mjs` runs each file in its own `bun test` process, so a hanging test cannot take down the rest of the lane. Pass `--jobs N` to parallelize or `--single-process` to share one process while debugging.

Launcher smoke artifacts go to `.tmp/launcher-smoke`. CI uploads them when the smoke step fails.

## Format

```bash
make pretty               # format the maintained source and stats files
bun run format:check:src  # check the same set without writing
```

`bun run format` runs Prettier over the whole repo. Use it only when you mean to.

## Config generation

```bash
bun run electron . --generate-config   # write a default config to ~/.config/SubMiner/config.jsonc (%APPDATA%\SubMiner\config.jsonc on Windows)
bun run generate:config-example        # regenerate config.example.jsonc from the defaults
```

`make generate-config` and `make generate-example-config` wrap the same commands.

Config definitions are split by domain under `src/config/definitions/`:

- defaults: `defaults-*.ts`
- option metadata: `options-*.ts`
- generated template sections and comments: `template-sections.ts`

`src/config/definitions.ts` composes them into the public API (`DEFAULT_CONFIG`, registries, template export). A new key also needs a resolver entry under `src/config/resolve/`, or the resolved config keeps the default.

## Documentation site

The user docs live in `docs-site/` (VitePress).

```bash
bun --cwd docs-site install
bun run docs:dev      # dev server at http://localhost:5173
bun run docs:test     # docs regression tests (links, pinned strings)
bun run docs:build    # production build into docs-site/.vitepress/dist
bun run docs:preview  # preview the build at http://localhost:4173
```

Run `bun run docs:test` and `bun run docs:build` whenever you change `docs-site/`.

Production uses the versioned build:

```bash
bun run docs:build:versioned
```

It writes `.tmp/docs-versioned-site`: the latest stable docs at `/` with a generated `/versions` page, and development docs at `/main/`. Prerelease tags are skipped. Stable archives under `/v/<version>/` are built once and stored in R2. Without R2 credentials the build skips archive sync, so a local run only produces `/` and `/main/`.

The `docs-pages` GitHub Actions workflow uploads that output to Cloudflare Pages with Wrangler. Cloudflare's Git-integration builds are disabled on purpose, so do not re-enable them in the dashboard. `docs-site/README.md` has the full deployment setup.

## Makefile targets

Run `make help` for the full list.

| Target                      | Description                                                  |
| --------------------------- | ------------------------------------------------------------ |
| `make deps`                 | Init submodules and install root, stats, and texthooker deps |
| `make build`                | Build the platform package for the current OS                |
| `make build-linux`          | Build the Linux package                                      |
| `make build-macos`          | Build the signed macOS package                               |
| `make build-macos-unsigned` | Build the unsigned macOS package                             |
| `make build-launcher`       | Generate the launcher in `dist/launcher/`                    |
| `make install`              | Install platform artifacts (wrapper, theme, AppImage or app) |
| `make pretty`               | Run scoped Prettier formatting                               |
| `make generate-config`      | Generate a default config                                    |

## Contributor notes

- See [Architecture](/architecture) for module boundaries and [IPC + runtime contracts](/ipc-contracts) before adding IPC channels.
- `src/core/services/overlay-manager.ts` owns overlay window and visibility state.
- In `src/main/` modules, pass simple dependencies as inline objects. Add a helper or adapter only when it adapts, validates, or gets reused.
- Packaged Linux desktop launches pass `--background` through `build.linux.executableArgs` in `package.json`.

## Environment variables

| Variable                           | Description                                                      |
| ---------------------------------- | ---------------------------------------------------------------- |
| `SUBMINER_APPIMAGE_PATH`           | SubMiner app binary the launcher uses for playback               |
| `SUBMINER_BINARY_PATH`             | Alias for `SUBMINER_APPIMAGE_PATH`                               |
| `SUBMINER_ROFI_THEME`              | rofi theme for the launcher picker                               |
| `SUBMINER_MPV_PLUGIN_PATH`         | mpv plugin directory the launcher injects                        |
| `SUBMINER_LOG_LEVEL`               | App log level (`debug`, `info`, `warn`, `error`)                 |
| `SUBMINER_MPV_LOG`                 | Shared mpv/app log file path                                     |
| `SUBMINER_JIMAKU_API_KEY`          | Jimaku API key for launcher subtitle downloads                   |
| `SUBMINER_JIMAKU_API_KEY_COMMAND`  | Command that prints the Jimaku API key                           |
| `SUBMINER_JIMAKU_API_BASE_URL`     | Jimaku API base URL                                              |
| `SUBMINER_JELLYFIN_ACCESS_TOKEN`   | Jellyfin access token, used before the stored encrypted session  |
| `SUBMINER_JELLYFIN_USER_ID`        | Jellyfin user ID                                                 |
| `SUBMINER_SKIP_MACOS_HELPER_BUILD` | Set to `1` to skip the macOS helper build during `bun run build` |
