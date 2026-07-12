import { fail } from '../log.js';
import { runAppCommandInteractive } from '../mpv.js';
import type { Args } from '../types.js';
import type { LauncherCommandContext } from './context.js';

export interface SyncCommandDeps {
  runAppCommand: (appPath: string, appArgs: string[]) => void;
  fail: (message: string) => never;
}

const defaultSyncCommandDeps: SyncCommandDeps = {
  runAppCommand: runAppCommandInteractive,
  fail,
};

type SyncArgs = Pick<
  Args,
  | 'syncHost'
  | 'syncSnapshotPath'
  | 'syncMergePath'
  | 'syncDirection'
  | 'syncRemoteCmd'
  | 'syncDbPath'
  | 'syncForce'
  | 'syncJson'
  | 'syncCheck'
  | 'syncMakeTemp'
  | 'syncRemoveTempPath'
  | 'logLevel'
>;

/** Rebuild the app's --sync-cli argv from the launcher's parsed sync args. */
export function buildSyncCliArgv(args: SyncArgs): string[] {
  const argv = ['--sync-cli', 'sync'];
  if (args.syncHost) argv.push(args.syncHost);
  if (args.syncSnapshotPath) argv.push('--snapshot', args.syncSnapshotPath);
  if (args.syncMergePath) argv.push('--merge', args.syncMergePath);
  if (args.syncMakeTemp) argv.push('--make-temp');
  if (args.syncRemoveTempPath) argv.push('--remove-temp', args.syncRemoveTempPath);
  if (args.syncDirection === 'push') argv.push('--push');
  if (args.syncDirection === 'pull') argv.push('--pull');
  if (args.syncCheck) argv.push('--check');
  if (args.syncRemoteCmd) argv.push('--remote-cmd', args.syncRemoteCmd);
  if (args.syncDbPath) argv.push('--db', args.syncDbPath);
  if (args.syncForce) argv.push('--force');
  if (args.syncJson) argv.push('--json');
  argv.push('--log-level', args.logLevel);
  return argv;
}

/**
 * `subminer sync` is a thin proxy: the sync engine only executes inside the
 * SubMiner app (--sync-cli mode, libsql), so the launcher and the app cannot
 * drift apart. The launcher contributes its parser/help and app discovery;
 * the child owns the terminal and its exit code becomes the launcher's.
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
  deps.runAppCommand(context.appPath, buildSyncCliArgv(context.args));
  return true;
}
