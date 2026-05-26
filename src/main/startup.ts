import { CliArgs } from '../cli/args';
import type { ResolvedConfig } from '../types';
import type { StartupBootstrapRuntimeDeps } from '../core/services/startup';
import type { LogLevelSource } from '../logger';

export interface StartupBootstrapRuntimeFactoryDeps {
  argv: string[];
  parseArgs: (argv: string[]) => CliArgs;
  setLogLevel: (level: string, source: LogLevelSource) => void;
  setLogRotation?: (rotation: number) => void;
  forceX11Backend: (args: CliArgs) => void;
  enforceUnsupportedWaylandMode: (args: CliArgs) => void;
  shouldStartApp: (args: CliArgs) => boolean;
  getDefaultSocketPath: () => string;
  defaultTexthookerPort: number;
  configDir: string;
  defaultConfig: ResolvedConfig;
  generateConfigTemplate: (config: ResolvedConfig) => string;
  generateDefaultConfigFile: (
    args: CliArgs,
    options: {
      configDir: string;
      defaultConfig: unknown;
      generateTemplate: (config: unknown) => string;
    },
  ) => Promise<number>;
  onConfigGenerated: (exitCode: number) => void;
  onGenerateConfigError: (error: Error) => void;
  startAppLifecycle: (args: CliArgs) => void;
}

export function createStartupBootstrapRuntimeDeps(
  params: StartupBootstrapRuntimeFactoryDeps,
): StartupBootstrapRuntimeDeps {
  return {
    argv: params.argv,
    parseArgs: params.parseArgs,
    setLogLevel: params.setLogLevel,
    setLogRotation: params.setLogRotation,
    forceX11Backend: (args: CliArgs) => params.forceX11Backend(args),
    enforceUnsupportedWaylandMode: (args: CliArgs) => params.enforceUnsupportedWaylandMode(args),
    getDefaultSocketPath: params.getDefaultSocketPath,
    defaultTexthookerPort: params.defaultTexthookerPort,
    runGenerateConfigFlow: (args: CliArgs) => {
      if (!args.generateConfig || params.shouldStartApp(args)) {
        return false;
      }
      params
        .generateDefaultConfigFile(args, {
          configDir: params.configDir,
          defaultConfig: params.defaultConfig,
          generateTemplate: (config: unknown) =>
            params.generateConfigTemplate(config as ResolvedConfig),
        })
        .then((exitCode) => {
          params.onConfigGenerated(exitCode);
        })
        .catch(params.onGenerateConfigError);
      return true;
    },
    startAppLifecycle: params.startAppLifecycle,
  };
}
