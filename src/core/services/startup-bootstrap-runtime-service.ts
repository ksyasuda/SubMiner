import { CliArgs } from "../../cli/args";

export interface StartupBootstrapRuntimeState {
  initialArgs: CliArgs;
  mpvSocketPath: string;
  texthookerPort: number;
  backendOverride: string | null;
  autoStartOverlay: boolean;
  texthookerOnlyMode: boolean;
}

export interface StartupBootstrapRuntimeDeps {
  argv: string[];
  parseArgs: (argv: string[]) => CliArgs;
  setLogLevelEnv: (level: string) => void;
  enableVerboseLogging: () => void;
  forceX11Backend: (args: CliArgs) => void;
  enforceUnsupportedWaylandMode: (args: CliArgs) => void;
  getDefaultSocketPath: () => string;
  defaultTexthookerPort: number;
  runGenerateConfigFlow: (args: CliArgs) => boolean;
  startAppLifecycle: (args: CliArgs) => void;
}

export function runStartupBootstrapRuntimeService(
  deps: StartupBootstrapRuntimeDeps,
): StartupBootstrapRuntimeState {
  const initialArgs = deps.parseArgs(deps.argv);

  if (initialArgs.logLevel) {
    deps.setLogLevelEnv(initialArgs.logLevel);
  } else if (initialArgs.verbose) {
    deps.enableVerboseLogging();
  }

  deps.forceX11Backend(initialArgs);
  deps.enforceUnsupportedWaylandMode(initialArgs);

  const state: StartupBootstrapRuntimeState = {
    initialArgs,
    mpvSocketPath: initialArgs.socketPath ?? deps.getDefaultSocketPath(),
    texthookerPort: initialArgs.texthookerPort ?? deps.defaultTexthookerPort,
    backendOverride: initialArgs.backend ?? null,
    autoStartOverlay: initialArgs.autoStartOverlay,
    texthookerOnlyMode: initialArgs.texthooker,
  };

  if (!deps.runGenerateConfigFlow(initialArgs)) {
    deps.startAppLifecycle(initialArgs);
  }

  return state;
}
