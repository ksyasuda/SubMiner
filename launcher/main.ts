import path from 'node:path';
import packageJson from '../package.json';
import { applyLogFileTogglesToEnv } from '../src/shared/log-files.js';
import {
  loadLauncherJellyfinConfig,
  loadLauncherLoggingConfig,
  loadLauncherMpvConfig,
  loadLauncherYoutubeSubgenConfig,
  parseArgs,
  readPluginRuntimeConfig,
} from './config.js';
import { fail, log } from './log.js';
import { findAppBinary, state } from './mpv.js';
import { nodeProcessAdapter } from './process-adapter.js';
import type { LauncherCommandContext } from './commands/context.js';
import { runDoctorCommand } from './commands/doctor-command.js';
import { runConfigCommand } from './commands/config-command.js';
import { runMpvPostAppCommand, runMpvPreAppCommand } from './commands/mpv-command.js';
import { runAppPassthroughCommand, runTexthookerCommand } from './commands/app-command.js';
import { runDictionaryCommand } from './commands/dictionary-command.js';
import { runLogsCommand } from './commands/logs-command.js';
import { runStatsCommand } from './commands/stats-command.js';
import { runJellyfinCommand } from './commands/jellyfin-command.js';
import { runPlaybackCommand } from './commands/playback-command.js';
import { runUpdateCommand } from './commands/update-command.js';

const APP_VERSION =
  typeof packageJson.version === 'string' && packageJson.version.trim()
    ? packageJson.version
    : 'unknown';

function createCommandContext(
  args: ReturnType<typeof parseArgs>,
  scriptPath: string,
  pluginRuntimeConfig: ReturnType<typeof readPluginRuntimeConfig>,
  appPath: string | null,
): LauncherCommandContext {
  return {
    args,
    scriptPath,
    scriptName: path.basename(scriptPath),
    mpvSocketPath: pluginRuntimeConfig.socketPath,
    pluginRuntimeConfig,
    appPath,
    launcherJellyfinConfig: loadLauncherJellyfinConfig(),
    processAdapter: nodeProcessAdapter,
  };
}

function ensureAppPath(context: LauncherCommandContext): string {
  if (context.appPath) {
    return context.appPath;
  }
  if (context.processAdapter.platform() === 'darwin') {
    fail(
      'SubMiner app binary not found. Install SubMiner.app to /Applications or ~/Applications, or set SUBMINER_APPIMAGE_PATH.',
    );
  }
  fail('SubMiner AppImage not found. Install to ~/.local/bin/ or set SUBMINER_APPIMAGE_PATH.');
}

async function main(): Promise<void> {
  const scriptPath = process.argv[1] || 'subminer';
  const scriptName = path.basename(scriptPath);
  const launcherConfig = loadLauncherYoutubeSubgenConfig();
  const launcherMpvConfig = loadLauncherMpvConfig();
  const launcherLoggingConfig = loadLauncherLoggingConfig();
  applyLogFileTogglesToEnv(launcherLoggingConfig.files);
  process.env.SUBMINER_LOG_ROTATION =
    launcherLoggingConfig.rotation !== undefined
      ? String(launcherLoggingConfig.rotation)
      : (process.env.SUBMINER_LOG_ROTATION ?? '7');
  const args = parseArgs(
    process.argv.slice(2),
    scriptName,
    launcherConfig,
    launcherMpvConfig,
    launcherLoggingConfig,
  );

  if (args.version) {
    console.log(`SubMiner ${APP_VERSION}`);
    return;
  }

  const pluginRuntimeConfig = readPluginRuntimeConfig(args.logLevel);
  const appPath = findAppBinary(scriptPath);

  log('debug', args.logLevel, `Wrapper log level set to: ${args.logLevel}`);

  const context = createCommandContext(args, scriptPath, pluginRuntimeConfig, appPath);

  if (runDoctorCommand(context)) {
    return;
  }

  if (runConfigCommand(context)) {
    return;
  }

  if (await runMpvPreAppCommand(context)) {
    return;
  }

  if (runLogsCommand(context)) {
    return;
  }

  const resolvedAppPath = ensureAppPath(context);
  state.appPath = resolvedAppPath;
  log('debug', args.logLevel, `Using SubMiner app binary: ${resolvedAppPath}`);
  const appContext: LauncherCommandContext = {
    ...context,
    appPath: resolvedAppPath,
  };

  if (runAppPassthroughCommand(appContext)) {
    return;
  }

  if (await runUpdateCommand(appContext)) {
    return;
  }

  if (await runMpvPostAppCommand(appContext)) {
    return;
  }

  if (runTexthookerCommand(appContext)) {
    return;
  }

  if (runDictionaryCommand(appContext)) {
    return;
  }

  if (await runStatsCommand(appContext)) {
    return;
  }

  if (await runJellyfinCommand(appContext)) {
    return;
  }

  await runPlaybackCommand(appContext);
}

main().catch((error: unknown) => {
  const message = error instanceof Error ? error.message : String(error);
  fail(message);
});
