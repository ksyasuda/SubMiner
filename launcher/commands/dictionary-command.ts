import { runAppCommandWithInherit } from '../mpv.js';
import { shouldForwardLogLevel } from '../types.js';
import type { LauncherCommandContext } from './context.js';

interface DictionaryCommandDeps {
  runAppCommandWithInherit: (appPath: string, appArgs: string[]) => void;
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

  const forwarded = [
    '--start',
    args.dictionaryCandidates
      ? '--dictionary-candidates'
      : args.dictionarySelect
        ? '--dictionary-select'
        : '--dictionary',
  ];
  if (args.dictionarySelect) {
    if (!args.dictionaryAnilistId) {
      throw new Error('Dictionary selection requires an AniList media ID.');
    }
    forwarded.push('--dictionary-anilist-id', String(args.dictionaryAnilistId));
  }
  if (typeof args.dictionaryTarget === 'string' && args.dictionaryTarget.trim()) {
    forwarded.push('--dictionary-target', args.dictionaryTarget);
  }
  if (shouldForwardLogLevel(args.logLevel)) {
    forwarded.push('--log-level', args.logLevel);
  }

  deps.runAppCommandWithInherit(appPath, forwarded);
  return true;
}
