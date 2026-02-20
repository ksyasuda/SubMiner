import type { CliArgs } from '../../cli/args';
import type { ResolvedConfig } from '../../types';
import type { LogLevelSource } from '../../logger';
import type { StartupBootstrapRuntimeFactoryDeps } from '../startup';

export function createBuildStartupBootstrapRuntimeFactoryDepsHandler(
  deps: StartupBootstrapRuntimeFactoryDeps,
) {
  return (): StartupBootstrapRuntimeFactoryDeps => ({
    argv: deps.argv,
    parseArgs: deps.parseArgs,
    setLogLevel: deps.setLogLevel,
    forceX11Backend: deps.forceX11Backend,
    enforceUnsupportedWaylandMode: deps.enforceUnsupportedWaylandMode,
    shouldStartApp: deps.shouldStartApp,
    getDefaultSocketPath: deps.getDefaultSocketPath,
    defaultTexthookerPort: deps.defaultTexthookerPort,
    configDir: deps.configDir,
    defaultConfig: deps.defaultConfig,
    generateConfigTemplate: deps.generateConfigTemplate,
    generateDefaultConfigFile: deps.generateDefaultConfigFile,
    onConfigGenerated: deps.onConfigGenerated,
    onGenerateConfigError: deps.onGenerateConfigError,
    startAppLifecycle: deps.startAppLifecycle,
  });
}

export type {
  CliArgs as StartupBuilderCliArgs,
  ResolvedConfig as StartupBuilderResolvedConfig,
  LogLevelSource as StartupBuilderLogLevelSource,
};
