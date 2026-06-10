import path from 'node:path';
import { loadRawConfigStrict } from './config/load';
import { resolveConfig } from './config/resolve';
import type { MpvLaunchMode, ResolvedConfig } from './types/config';
import type { LogFileToggles, LogRotation } from './shared/log-files';
import type { SharedLogLevel } from './shared/mpv-logging-args';
import type { SubminerPluginRuntimeScriptOptConfig } from './shared/subminer-plugin-script-opts';

export interface ConfiguredWindowsMpvLaunch {
  executablePath: string;
  launchMode: MpvLaunchMode;
  logLevel: SharedLogLevel;
  logRotation: LogRotation;
  logFiles: LogFileToggles;
  pluginRuntimeConfig: SubminerPluginRuntimeScriptOptConfig;
}

export function buildWindowsMpvPluginRuntimeConfig(
  config: Pick<ResolvedConfig, 'auto_start_overlay' | 'logging' | 'mpv' | 'texthooker'>,
): SubminerPluginRuntimeScriptOptConfig {
  return {
    socketPath: config.mpv.socketPath,
    binaryPath: config.mpv.subminerBinaryPath,
    backend: config.mpv.backend,
    logLevel: config.logging.level,
    logRotation: config.logging.rotation,
    autoStart: config.mpv.autoStartSubMiner,
    autoStartVisibleOverlay: config.auto_start_overlay,
    autoStartPauseUntilReady: config.mpv.pauseUntilOverlayReady,
    texthookerEnabled: config.texthooker.launchAtStartup,
  };
}

export function readConfiguredWindowsMpvLaunch(configDir: string): ConfiguredWindowsMpvLaunch {
  const loadResult = loadRawConfigStrict({
    configDir,
    configFileJsonc: path.join(configDir, 'config.jsonc'),
    configFileJson: path.join(configDir, 'config.json'),
  });
  const rawConfig = loadResult.ok ? loadResult.config : {};
  const { resolved } = resolveConfig(rawConfig);

  return {
    executablePath: resolved.mpv.executablePath,
    launchMode: resolved.mpv.launchMode,
    logLevel: resolved.logging.level,
    logRotation: resolved.logging.rotation,
    logFiles: resolved.logging.files,
    pluginRuntimeConfig: buildWindowsMpvPluginRuntimeConfig(resolved),
  };
}
