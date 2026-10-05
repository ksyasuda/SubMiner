<!-- read_when: choosing what tests/build steps to run before handoff -->

# Verification

Status: active  
Last verified: 2026-10-05
Owner: Kyle Yasuda  
Read when: selecting the right verification lane for a change

## Lane Infrastructure

- Lane membership is defined once in `scripts/test-lanes.ts` and discovered by
  directory, so new test files join their lane automatically; never hand-list test
  files in `package.json`.
- `scripts/run-test-lane.mjs` runs each test file in its own `bun test` process
  (per-file isolation with a wall timeout) so a hanging test or leaked global in
  one file cannot cascade into the rest of the lane. Files run in parallel, one
  worker per CPU by default; `--jobs 1` runs serially and `--single-process`
  restores the shared-process mode for debugging.
- `bun run test:fast` is the full source gate: discovered `src/**`, launcher
  unit, and `scripts/**`.
- `.github/workflows/quality-gate.yml` is the reusable `workflow_call` gate for
  pull requests, stable tags, and prerelease tags. Keep common quality steps
  there instead of copying them into caller workflows.
- The reusable gate installs Lua and runs `bun run test:env`, so the shipped mpv
  plugin tests and launcher smoke run for every pull request and tagged release.
  Lua installation uses only the runner's Ubuntu package sources so unrelated
  third-party repository failures do not block the gate.
- In the reusable gate, `test:coverage:src` is also the blocking execution of the
  discovered `src/**` test lane. The coverage runner returns the failing test's
  status, so CI does not rerun that lane through `test:fast`. Launcher unit and
  script tests still run separately because they are outside the coverage lane.
- Launcher smoke artifacts are uploaded after `test:env` fails. CI does not rerun
  launcher smoke solely to collect the same artifacts.

## Default Handoff Gate

```bash
bun run typecheck
bun run test:fast
bun run test:env
bun run build
bun run test:smoke:dist
```

If `docs-site/` changed, also run:

```bash
bun run docs:test
bun run docs:build
```

## Cheap-First Lane Selection

- User-facing `docs-site/` changes: `bun run docs:test`, `bun run docs:build`
- Internal KB, `AGENTS.md`, or `.agents/skills/**` changes: `bun run test:docs:kb`
- Config/schema/defaults: `bun run test:config`, then `bun run generate:config-example` if template/defaults changed
- Launcher/plugin: `bun run test:launcher` or `bun run test:env`
- Runtime-compat / compiled behavior after `bun run build`: `bun run test:runtime:compat`
- Stats dashboard UI: `bun run test:stats`
- Overlay, renderer, mpv, or mining behavior that unit tests cannot show: `bun run build`,
  then `bun run test:e2e` (see [Headless End-to-End Harness](#headless-end-to-end-harness))
- Build/release scripts (`scripts/**`): `bun run test:scripts`
- Packaging: build the platform package, then run `bun run test:package <resources-directory>`.
  On headless Linux: `xvfb-run -a bun run test:package release/linux-unpacked/resources`.
  Content checks run inside the electron-builder afterPack hook. See the
  [release guide](../RELEASING.md#package-contents-checks) for the
  installed-app verification checklist.
- Coverage for the maintained source lane: `bun run test:coverage:src`
- Deep/local full gate: default handoff gate above

## Coverage Reporting

- `bun run test:coverage:src` runs the same discovered `bun-src-full` membership as
  `test:src` through a sharded coverage runner: one Bun coverage process per test
  file (parallel, one per CPU unless `--jobs N`), then merged LCOV output.
- A failing coverage shard stops the runner with a nonzero status. Coverage is a
  source test gate, not a report-only step.
- Machine-readable output lands at `coverage/test-src/lcov.info`.
- Every reusable quality-gate run uploads that LCOV file as the
  `coverage-test-src` artifact.

## Compiled Runtime Smoke

- `bun run test:smoke:dist` and its `test:runtime:compat` alias require an existing
  full build and fail with the missing artifact paths when `dist/` or the stats UI
  bundle is absent.
- The check runs the emitted stats daemon under Electron's Node runtime. It opens
  the production HTTP server, queries the overview endpoint through native
  libsql-backed storage, and leaves an HTTP request body unfinished before
  shutting the daemon down. It verifies a clean exit and that the port and
  ownership state are released despite the unfinished request.
- The check also occupies the configured port, requires startup to fail without
  stale ownership state, releases the conflict, and verifies a clean retry.
- This is not a full Electron UI startup check. It does not require a display and
  makes no claims about renderer, window, tray, or mpv behavior.

## Headless End-to-End Harness

`bun run test:e2e` boots the built app against a real mpv and asserts on what the
overlay renders and what reaches AnkiConnect. Use it instead of a live desktop test
whenever the behavior does not depend on a specific compositor.

- Needs an existing `bun run build`, plus `mpv` and `ffmpeg` on `PATH` (and `Xvfb` on
  Linux). It is not part of `test:fast` or the handoff gate. The `e2e` job in
  `.github/workflows/ci.yml` runs it on Linux, macOS, and Windows for pull requests and
  `main`, outside the reusable quality gate, and uploads the session logs when it fails.
- Every run is sandboxed in a temp dir: its own config and Electron profile, mpv IPC
  socket, and a fake AnkiConnect on a free port. It never touches the user's running
  SubMiner, Anki, or stats database.
- Media is synthetic: an ffmpeg test-pattern clip, a three-cue Japanese `.srt`, and a
  tiny Yomitan dictionary covering those cues (`e2e/harness/fixtures.ts`). The harness
  imports the dictionary on a first boot, then restarts the app for the scenario.
- Display: Linux runs everything inside a private Xvfb server, so no window reaches the
  desktop. macOS and Windows have no offscreen display server, so windows open in the
  current session there; `SUBMINER_E2E_DISPLAY=host` forces the same on Linux. The
  macOS and Windows paths are written but have not been run yet.
- The app is driven over the Chrome DevTools Protocol (DOM queries, input events,
  per-window screenshots), mpv over its JSON IPC socket, and app commands through the
  normal CLI handoff (`--mine-sentence`, `--toggle-subtitle-sidebar`, ...).
- Scenarios live in `e2e/*.e2e.test.ts` and share one session per file through
  `startE2eSession()`. Pass `config` to override settings for a scenario file.
- On a failed startup the temp dir is kept and its path is in the error. Set
  `SUBMINER_E2E_KEEP=1` to keep it after any run (`app-stdio.log`, `app.log`, `mpv.log`).

For exploratory checks, keep one session up and poke at it:

```bash
bun run e2e start &                 # prints the session JSON, stays up until stopped
bun run e2e mpv seek 2 absolute+exact
bun run e2e eval overlay 'document.getElementById("subtitleRoot").textContent'
bun run e2e shot overlay /tmp/overlay.png
bun run e2e app --mine-sentence
bun run e2e anki                    # AnkiConnect requests received so far
bun run e2e stop
```

Out of scope: compositor-specific behavior (Hyprland placement, stacking, fullscreen,
click-through), real GPU video output, and global shortcuts. Xvfb runs without a window
manager or compositor, so the transparent overlay is not composited over the video;
check overlay visuals with per-window screenshots, and use the live desktop for the rest.

## Dependency Audit Policy

- `bun audit --audit-level high` blocks the reusable quality gate.
- Keep security overrides and dependency patches at the minimum fixed version.
  Remove them after the owning package ships and adopts a compatible fix.

## Rules

- Capture exact failing command and error when verification breaks.
- Prefer the cheapest sufficient lane first.
- Escalate when the change crosses boundaries or touches release-sensitive behavior.
- Never hand-edit `dist/launcher/subminer`; validate it through build/test flow instead.
