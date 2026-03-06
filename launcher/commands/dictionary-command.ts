import { runAppCommandWithInherit } from '../mpv.js';
import type { LauncherCommandContext } from './context.js';

interface DictionaryCommandDeps {
  runAppCommandWithInherit: (appPath: string, appArgs: string[]) => never;
}

const defaultDeps: DictionaryCommandDeps = {
  runAppCommandWithInherit,
};

export function runDictionaryCommand(
  context: LauncherCommandContext,
  deps: DictionaryCommandDeps = defaultDeps,
): boolean {
  const { args, appPath } = context;
  if (!args.dictionary || !appPath) {
    return false;
  }

  const forwarded = ['--dictionary'];
  if (typeof args.dictionaryTarget === 'string' && args.dictionaryTarget.trim()) {
    forwarded.push('--dictionary-target', args.dictionaryTarget);
  }
  if (args.logLevel !== 'info') {
    forwarded.push('--log-level', args.logLevel);
  }

  deps.runAppCommandWithInherit(appPath, forwarded);
  return true;
}
