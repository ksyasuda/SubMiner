import { SYNC_CLI_FLAG } from '../../src/core/services/stats-sync/cli-args.js';
import { fail } from '../log.js';
import { runAppCommandInteractive } from '../mpv.js';
import type { LauncherCommandContext } from './context.js';

export interface SyncCommandDeps {
  runAppCommand: (appPath: string, appArgs: string[]) => void;
  fail: (message: string) => never;
}

const defaultSyncCommandDeps: SyncCommandDeps = {
  runAppCommand: runAppCommandInteractive,
  fail,
};

/**
 * `subminer sync` is a thin proxy: the sync engine only executes inside the
 * SubMiner app (--sync-cli mode, libsql). The launcher contributes its
 * parser/help and app discovery; the app's parseSyncCliTokens owns validation,
 * so its errors reach the terminal through the child's inherited stdio. The
 * child owns the terminal and its exit code becomes the launcher's.
 */
export async function runSyncCommand(
  context: LauncherCommandContext,
  inputDeps: Partial<SyncCommandDeps> = {},
): Promise<boolean> {
  const deps = { ...defaultSyncCommandDeps, ...inputDeps };
  if (!context.args.sync) return false;
  if (!context.appPath) {
    deps.fail(
      context.processAdapter.platform() === 'darwin'
        ? 'SubMiner app binary not found (sync runs inside the app). Install SubMiner.app to /Applications or ~/Applications, or set SUBMINER_APPIMAGE_PATH.'
        : 'SubMiner app binary not found (sync runs inside the app). Install the SubMiner app, or set SUBMINER_APPIMAGE_PATH.',
    );
    return true; // fail() never returns; this only satisfies control-flow analysis
  }
  deps.runAppCommand(context.appPath, [
    SYNC_CLI_FLAG,
    'sync',
    ...context.args.syncCliTokens,
    '--log-level',
    context.args.logLevel,
  ]);
  return true;
}
