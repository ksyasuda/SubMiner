import { fail } from './log.js';
import type {
  Args,
  LauncherJellyfinConfig,
  LauncherYoutubeSubgenConfig,
  LogLevel,
  PluginRuntimeConfig,
} from './types.js';
import {
  applyInvocationsToArgs,
  applyRootOptionsToArgs,
  createDefaultArgs,
} from './config/args-normalizer.js';
import { parseCliPrograms, resolveTopLevelCommand } from './config/cli-parser-builder.js';
import { parseLauncherJellyfinConfig } from './config/jellyfin-config.js';
import { readPluginRuntimeConfig as readPluginRuntimeConfigValue } from './config/plugin-runtime-config.js';
import { readLauncherMainConfigObject } from './config/shared-config-reader.js';
import { parseLauncherYoutubeSubgenConfig } from './config/youtube-subgen-config.js';

export function loadLauncherYoutubeSubgenConfig(): LauncherYoutubeSubgenConfig {
  const root = readLauncherMainConfigObject();
  if (!root) return {};
  return parseLauncherYoutubeSubgenConfig(root);
}

export function loadLauncherJellyfinConfig(): LauncherJellyfinConfig {
  const root = readLauncherMainConfigObject();
  if (!root) return {};
  return parseLauncherJellyfinConfig(root);
}

export function readPluginRuntimeConfig(logLevel: LogLevel): PluginRuntimeConfig {
  return readPluginRuntimeConfigValue(logLevel);
}

export function parseArgs(
  argv: string[],
  scriptName: string,
  launcherConfig: LauncherYoutubeSubgenConfig,
): Args {
  const topLevelCommand = resolveTopLevelCommand(argv);
  const parsed = createDefaultArgs(launcherConfig);

  if (topLevelCommand && (topLevelCommand.name === 'app' || topLevelCommand.name === 'bin')) {
    parsed.appPassthrough = true;
    parsed.appArgs = argv.slice(topLevelCommand.index + 1);
    return parsed;
  }

  let cliResult: ReturnType<typeof parseCliPrograms>;
  try {
    cliResult = parseCliPrograms(argv, scriptName);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    fail(message);
  }

  applyRootOptionsToArgs(parsed, cliResult.options, cliResult.rootTarget);
  applyInvocationsToArgs(parsed, cliResult.invocations);
  return parsed;
}
