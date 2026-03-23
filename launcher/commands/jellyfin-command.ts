import { fail } from '../log.js';
import { runAppCommandWithInherit } from '../mpv.js';
import { commandExists } from '../util.js';
import { runJellyfinPlayMenu } from '../jellyfin.js';
import type { LauncherCommandContext } from './context.js';

export async function runJellyfinCommand(context: LauncherCommandContext): Promise<boolean> {
  const { args, appPath, scriptPath, mpvSocketPath, launcherJellyfinConfig } = context;
  if (!appPath) {
    return false;
  }

  const appendPasswordStore = (forwarded: string[]): void => {
    if (args.passwordStore) {
      forwarded.push('--password-store', args.passwordStore);
    }
  };

  if (args.jellyfin) {
    const forwarded = ['--jellyfin'];
    if (args.logLevel !== 'info') forwarded.push('--log-level', args.logLevel);
    appendPasswordStore(forwarded);
    runAppCommandWithInherit(appPath, forwarded);
    return true;
  }

  if (args.jellyfinLogin) {
    const serverUrl = args.jellyfinServer || launcherJellyfinConfig.serverUrl || '';
    const username = args.jellyfinUsername || launcherJellyfinConfig.username || '';
    const password = args.jellyfinPassword || '';
    if (!serverUrl || !username || !password) {
      fail(
        '--jellyfin-login requires server, username, and password. Pass flags or run `subminer --jellyfin`.',
      );
    }
    const forwarded = [
      '--jellyfin-login',
      '--jellyfin-server',
      serverUrl,
      '--jellyfin-username',
      username,
      '--jellyfin-password',
      password,
    ];
    if (args.logLevel !== 'info') forwarded.push('--log-level', args.logLevel);
    appendPasswordStore(forwarded);
    runAppCommandWithInherit(appPath, forwarded);
    return true;
  }

  if (args.jellyfinLogout) {
    const forwarded = ['--jellyfin-logout'];
    if (args.logLevel !== 'info') forwarded.push('--log-level', args.logLevel);
    appendPasswordStore(forwarded);
    runAppCommandWithInherit(appPath, forwarded);
    return true;
  }

  if (args.jellyfinPlay) {
    if (!args.useRofi && !commandExists('fzf')) {
      fail('fzf not found. Install fzf or use -R for rofi.');
    }
    if (args.useRofi && !commandExists('rofi')) {
      fail('rofi not found. Install rofi or omit -R for fzf.');
    }
    await runJellyfinPlayMenu(appPath, args, scriptPath, mpvSocketPath);
    return true;
  }

  if (args.jellyfinDiscovery) {
    const forwarded = ['--background', '--jellyfin-remote-announce'];
    if (args.logLevel !== 'info') forwarded.push('--log-level', args.logLevel);
    appendPasswordStore(forwarded);
    runAppCommandWithInherit(appPath, forwarded);
    return true;
  }

  return Boolean(
    args.jellyfin ||
    args.jellyfinLogin ||
    args.jellyfinLogout ||
    args.jellyfinPlay ||
    args.jellyfinDiscovery,
  );
}
