import path from 'node:path';
import { loadRawConfigStrict } from './config/load';
import { normalizeLaunchMpvExtraArgs, normalizeLaunchMpvTargets } from './main-entry-runtime';
import type { WindowsMpvLaunchDeps } from './main/runtime/windows-mpv-launch';
import { resolveRawMpvLaunchMode } from './shared/mpv-launch-mode';
import type { MpvLaunchMode } from './types/config';

export interface ConfiguredWindowsMpvLaunchConfig {
  executablePath: string;
  launchMode: MpvLaunchMode;
}

export function readConfiguredWindowsMpvLaunchConfig(
  configDir: string,
): ConfiguredWindowsMpvLaunchConfig {
  const loadResult = loadRawConfigStrict({
    configDir,
    configFileJsonc: path.join(configDir, 'config.jsonc'),
    configFileJson: path.join(configDir, 'config.json'),
  });
  if (!loadResult.ok) {
    return {
      executablePath: '',
      launchMode: 'normal',
    };
  }

  return {
    executablePath:
      typeof loadResult.config.mpv?.executablePath === 'string'
        ? loadResult.config.mpv.executablePath.trim()
        : '',
    launchMode: resolveRawMpvLaunchMode(loadResult.config.mpv) ?? 'normal',
  };
}

export async function handleLaunchMpvEntry(options: {
  argv: string[];
  userDataPath: string;
  processExecPath: string;
  createWindowsMpvLaunchDeps: () => WindowsMpvLaunchDeps;
  launchWindowsMpv: (
    targets: string[],
    deps: WindowsMpvLaunchDeps,
    extraArgs: string[],
    binaryPath?: string,
    pluginEntrypointPath?: string,
    configuredMpvPath?: string,
    launchMode?: MpvLaunchMode,
  ) => Promise<{ ok: boolean; mpvPath: string }>;
  resolveBundledWindowsMpvPluginEntrypoint: () => string | undefined;
}): Promise<{ ok: boolean; mpvPath: string }> {
  const configuredMpvLaunch = readConfiguredWindowsMpvLaunchConfig(options.userDataPath);
  return await options.launchWindowsMpv(
    normalizeLaunchMpvTargets(options.argv),
    options.createWindowsMpvLaunchDeps(),
    normalizeLaunchMpvExtraArgs(options.argv),
    options.processExecPath,
    options.resolveBundledWindowsMpvPluginEntrypoint(),
    configuredMpvLaunch.executablePath,
    configuredMpvLaunch.launchMode,
  );
}
