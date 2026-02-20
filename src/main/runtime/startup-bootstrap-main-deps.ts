import type { CliArgs } from '../../cli/args';
import type { LogLevelSource } from '../../logger';
import type { ResolvedConfig } from '../../types';
import type { StartupBootstrapRuntimeFactoryDeps } from '../startup';

export function createBuildStartupBootstrapMainDepsHandler(deps: {
  argv: string[];
  parseArgs: (argv: string[]) => CliArgs;
  setLogLevel: (level: string, source: LogLevelSource) => void;
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
  setExitCode: (code: number) => void;
  quitApp: () => void;
  logGenerateConfigError: (message: string) => void;
  startAppLifecycle: (args: CliArgs) => void;
}) {
  return (): StartupBootstrapRuntimeFactoryDeps => ({
    argv: deps.argv,
    parseArgs: (argv: string[]) => deps.parseArgs(argv),
    setLogLevel: (level: string, source: LogLevelSource) => deps.setLogLevel(level, source),
    forceX11Backend: (args: CliArgs) => deps.forceX11Backend(args),
    enforceUnsupportedWaylandMode: (args: CliArgs) => deps.enforceUnsupportedWaylandMode(args),
    shouldStartApp: (args: CliArgs) => deps.shouldStartApp(args),
    getDefaultSocketPath: () => deps.getDefaultSocketPath(),
    defaultTexthookerPort: deps.defaultTexthookerPort,
    configDir: deps.configDir,
    defaultConfig: deps.defaultConfig,
    generateConfigTemplate: (config: ResolvedConfig) => deps.generateConfigTemplate(config),
    generateDefaultConfigFile: (
      args: CliArgs,
      options: {
        configDir: string;
        defaultConfig: unknown;
        generateTemplate: (config: unknown) => string;
      },
    ) => deps.generateDefaultConfigFile(args, options),
    onConfigGenerated: (exitCode: number) => {
      deps.setExitCode(exitCode);
      deps.quitApp();
    },
    onGenerateConfigError: (error: Error) => {
      deps.logGenerateConfigError(`Failed to generate config: ${error.message}`);
      deps.setExitCode(1);
      deps.quitApp();
    },
    startAppLifecycle: (args: CliArgs) => deps.startAppLifecycle(args),
  });
}
