import { fail, log } from '../log.js';
import {
  waitForUnixSocketReady,
  launchMpvIdleDetached,
  resolveLauncherRuntimePluginPath,
} from '../mpv.js';
import type { LauncherCommandContext } from './context.js';

interface MpvCommandDeps {
  waitForUnixSocketReady(socketPath: string, timeoutMs: number): Promise<boolean>;
  launchMpvIdleDetached(
    socketPath: string,
    appPath: string,
    args: LauncherCommandContext['args'],
    runtimePluginPath?: string | null,
    runtimePluginConfig?: LauncherCommandContext['pluginRuntimeConfig'],
  ): Promise<void>;
}

const defaultDeps: MpvCommandDeps = {
  waitForUnixSocketReady,
  launchMpvIdleDetached,
};

export async function runMpvPreAppCommand(
  context: LauncherCommandContext,
  deps: MpvCommandDeps = defaultDeps,
): Promise<boolean> {
  const { args, mpvSocketPath, processAdapter } = context;
  if (args.mpvSocket) {
    processAdapter.writeStdout(`${mpvSocketPath}\n`);
    return true;
  }

  if (!args.mpvStatus) {
    return false;
  }

  const ready = await deps.waitForUnixSocketReady(mpvSocketPath, 500);
  log(
    ready ? 'info' : 'warn',
    args.logLevel,
    `[mpv] socket ${ready ? 'ready' : 'not ready'}: ${mpvSocketPath}`,
  );
  processAdapter.exit(ready ? 0 : 1);
  return true;
}

export async function runMpvPostAppCommand(
  context: LauncherCommandContext,
  deps: MpvCommandDeps = defaultDeps,
): Promise<boolean> {
  const { args, appPath, scriptPath, mpvSocketPath, pluginRuntimeConfig } = context;
  if (!args.mpvIdle) {
    return false;
  }
  if (!appPath) {
    fail('SubMiner app binary not found. Install to ~/.local/bin/ or set SUBMINER_APPIMAGE_PATH.');
  }

  await deps.launchMpvIdleDetached(
    mpvSocketPath,
    appPath,
    args,
    resolveLauncherRuntimePluginPath({ appPath, scriptPath }),
    {
      ...pluginRuntimeConfig,
      backend: args.backend,
      texthookerEnabled: args.useTexthooker && pluginRuntimeConfig.texthookerEnabled,
    },
  );
  const ready = await deps.waitForUnixSocketReady(mpvSocketPath, 8000);
  if (!ready) {
    fail(`MPV IPC socket not ready after idle launch: ${mpvSocketPath}`);
  }
  log('info', args.logLevel, `[mpv] idle instance ready on ${mpvSocketPath}`);
  return true;
}
