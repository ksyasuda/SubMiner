# Architecture

SubMiner uses a service-oriented Electron architecture with a composition-oriented main process and a modular renderer process.

## Goals

- Keep behavior stable while reducing coupling.
- Prefer small, single-purpose units that can be tested in isolation.
- Keep `main.ts` focused on wiring and state ownership, not implementation detail.
- Follow Unix-style composability:
  - each service does one job
  - services compose through explicit inputs/outputs
  - orchestration is separate from implementation

## Project Structure

```text
src/
  main.ts                  # Composition root — lifecycle wiring and state ownership
  preload.ts               # Electron preload bridge
  types.ts                 # Shared type definitions
  core/
    services/              # ~35 focused service modules (see below)
    utils/                 # Pure helpers and coercion/config utilities
  cli/                     # CLI parsing and help output
  config/                  # Config schema, defaults, validation, template generation
  renderer/                # Overlay renderer (modularized UI/runtime)
  window-trackers/         # Backend-specific tracker implementations (Hyprland, X11, macOS)
  jimaku/                  # Jimaku API integration helpers
  subsync/                 # Subtitle sync (alass/ffsubsync) helpers
  subtitle/                # Subtitle processing utilities
  tokenizers/              # Tokenizer implementations
  token-mergers/           # Token merge strategies
  translators/             # AI translation providers
```

### Service Layer (`src/core/services/`)

- **Startup** — `startup-service`, `app-lifecycle-service`
- **Overlay** — `overlay-manager-service`, `overlay-window-service`, `overlay-visibility-service`, `overlay-bridge-service`, `overlay-runtime-init-service`
- **Shortcuts** — `shortcut-service`, `overlay-shortcut-service`, `overlay-shortcut-handler`, `shortcut-fallback-service`, `numeric-shortcut-service`
- **MPV** — `mpv-service`, `mpv-control-service`, `mpv-render-metrics-service`
- **IPC** — `ipc-service`, `ipc-command-service`, `runtime-options-ipc-service`
- **Mining** — `mining-service`, `field-grouping-service`, `field-grouping-overlay-service`, `anki-jimaku-service`, `anki-jimaku-ipc-service`
- **Subtitles** — `subtitle-ws-service`, `subtitle-position-service`, `secondary-subtitle-service`, `tokenizer-service`
- **Integrations** — `jimaku-service`, `subsync-service`, `subsync-runner-service`, `texthooker-service`, `yomitan-extension-loader-service`, `yomitan-settings-service`
- **Config** — `runtime-config-service`, `cli-command-service`

### Renderer Layer (`src/renderer/`)

The overlay renderer is split by concern so `renderer.ts` stays focused on bootstrapping, IPC wiring, and module composition.

```text
src/renderer/
  renderer.ts              # Entrypoint/orchestration only
  context.ts               # Shared runtime context contract
  state.ts                 # Centralized renderer mutable state
  subtitle-render.ts       # Primary/secondary subtitle rendering + style application
  positioning.ts           # Visible/invisible positioning + mpv metrics layout
  handlers/
    keyboard.ts            # Keybindings, chord handling, modal key routing
    mouse.ts               # Hover/drag behavior, selection + observer wiring
  modals/
    jimaku.ts              # Jimaku modal flow
    kiku.ts                # Kiku field-grouping modal flow
    runtime-options.ts     # Runtime options modal flow
    subsync.ts             # Manual subsync modal flow
  utils/
    dom.ts                 # Required DOM lookups + typed handles
    platform.ts            # Layer/platform capability detection
```

## Flow Diagram

```mermaid
flowchart TD
  classDef root fill:#1f2937,stroke:#111827,color:#f9fafb,stroke-width:1.5px;
  classDef orchestration fill:#334155,stroke:#0f172a,color:#e2e8f0;
  classDef domain fill:#1d4ed8,stroke:#1e3a8a,color:#dbeafe;
  classDef boundary fill:#065f46,stroke:#064e3b,color:#d1fae5;

  subgraph Entry["Entrypoint"]
    Main["src/main.ts\ncomposition root"]
  end
  class Main root;

  subgraph Boot["Startup Orchestration"]
    Startup["startup-service"]
    Lifecycle["app-lifecycle-service"]
    AppReady["app-ready flow"]
  end
  class Startup,Lifecycle,AppReady orchestration;

  subgraph Runtime["Runtime Domains"]
    OverlayMgr["overlay-manager-service"]
    OverlayWindow["overlay-window-service"]
    OverlayVisibility["overlay-visibility-service"]
    Ipc["ipc-service\nipc-command-service"]
    RuntimeOpts["runtime-options-ipc-service"]
    Mpv["mpv-service\nmpv-control-service"]
    Subtitle["subtitle-ws-service\nsecondary-subtitle-service"]
    Shortcuts["shortcut-service\noverlay-shortcut-service"]
  end
  class OverlayMgr,OverlayWindow,OverlayVisibility,Ipc,RuntimeOpts,Mpv,Subtitle,Shortcuts domain;

  subgraph Adapters["External Boundaries"]
    Config["src/config/*"]
    Cli["src/cli/*"]
    Trackers["src/window-trackers/*"]
    Integrations["src/jimaku/*\nsrc/subsync/*"]
  end
  class Config,Cli,Trackers,Integrations boundary;

  Main -->|bootstraps| Startup
  Main -->|registers lifecycle hooks| Lifecycle
  Lifecycle -->|triggers| AppReady

  Main -->|wires| OverlayMgr
  Main -->|wires| Ipc
  Main -->|wires| Mpv
  Main -->|wires| Shortcuts
  Main -->|wires| RuntimeOpts
  Main -->|wires| Subtitle

  Main -->|loads| Config
  Main -->|parses| Cli
  Main -->|delegates backend state| Trackers
  Main -->|calls integrations| Integrations

  OverlayMgr -->|creates window| OverlayWindow
  OverlayMgr -->|applies visibility policy| OverlayVisibility
  Ipc -->|updates| RuntimeOpts
  Mpv -->|feeds timing + subtitle context| Subtitle
  Shortcuts -->|drives overlay actions| OverlayMgr
```

## Composition Pattern

Most runtime code follows a dependency-injection pattern:

1. Define a service interface in `src/core/services/*`.
2. Keep core logic in pure or side-effect-bounded functions.
3. Build runtime deps in `main.ts`; extract an adapter/helper only when it adds meaningful behavior or reuse.
4. Call the service from lifecycle/command wiring points.

This keeps side effects explicit and makes behavior easy to unit-test with fakes.

## Lifecycle Model

- **Startup:**
  - `startup-service` handles initial argv/env/backend setup and decides generate-config flow vs app lifecycle start.
  - `app-lifecycle-service` handles Electron single-instance + lifecycle event registration.
  - App-ready flow performs ready-time initialization (config load, websocket policy, tokenizer/tracker setup, overlay auto-init decisions).
- **Runtime:**
  - CLI/shortcut/IPC events map to service calls.
  - Overlay and MPV state sync through dedicated services.
  - Runtime options and mining flows are coordinated via service boundaries.
- **Shutdown:**
  - `app-lifecycle-service` registers cleanup hooks (`will-quit`) while teardown behavior stays delegated to focused services from `main.ts`.

```mermaid
flowchart TD
  classDef phase fill:#334155,stroke:#0f172a,color:#e2e8f0;
  classDef decision fill:#7c2d12,stroke:#431407,color:#ffedd5;
  classDef runtime fill:#0369a1,stroke:#0c4a6e,color:#e0f2fe;
  classDef shutdown fill:#14532d,stroke:#052e16,color:#dcfce7;

  Args["CLI args / env"] --> Startup["startup-service"]
  class Args,Startup phase;

  Startup --> Decision{"generate-config?"}
  class Decision decision;

  Decision -->|yes| WriteConfig["write config + exit"]
  Decision -->|no| Lifecycle["app-lifecycle-service"]
  class WriteConfig,Lifecycle phase;

  Lifecycle --> Ready["app-ready flow\n(config + websocket policy + tracker/tokenizer init)"]
  class Ready phase;

  Ready --> RuntimeBus["event loop:\nIPC + shortcuts + mpv events"]
  RuntimeBus --> Overlay["overlay visibility + mining actions"]
  RuntimeBus --> Subtitle["subtitle + secondary-subtitle processing"]
  RuntimeBus --> Subsync["subsync / jimaku integration actions"]
  class RuntimeBus,Overlay,Subtitle,Subsync runtime;

  RuntimeBus --> WillQuit["Electron will-quit"]
  WillQuit --> Cleanup["service-level teardown\n(unregister hooks, close resources)"]
  class WillQuit,Cleanup shutdown;
```

## Why This Design

- **Smaller blast radius:** changing one feature usually touches one service.
- **Better testability:** most behavior can be tested without Electron windows/mpv.
- **Better reviewability:** PRs can be scoped to one subsystem.
- **Backward compatibility:** CLI flags and IPC channels can remain stable while internals evolve.

## Extension Rules

- Add behavior to an existing service or a new `src/core/services/*` file, not as ad-hoc logic in `main.ts`.
- Keep service APIs explicit and narrowly scoped.
- Prefer additive changes that preserve existing CLI flags and IPC channel behavior.
- Add/update unit tests for each service extraction or behavior change.
- For cross-cutting changes, extract-first then refactor internals after parity is verified.
