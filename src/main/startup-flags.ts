import type { CliArgs } from '../cli/args';
import { isStandaloneTexthookerCommand, shouldRunSettingsOnlyStartup } from '../cli/args';

export function getPasswordStoreArg(argv: string[]): string | null {
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg?.startsWith('--password-store')) {
      continue;
    }

    if (arg === '--password-store') {
      const value = argv[i + 1];
      if (value && !value.startsWith('--')) {
        return value;
      }
      return null;
    }

    const [prefix, value] = arg.split('=', 2);
    if (prefix === '--password-store' && value && value.trim().length > 0) {
      return value.trim();
    }
  }
  return null;
}

export function normalizePasswordStoreArg(value: string): string {
  const normalized = value.trim();
  if (normalized.toLowerCase() === 'gnome') {
    return 'gnome-libsecret';
  }
  return normalized;
}

export function getDefaultPasswordStore(): string {
  return 'gnome-libsecret';
}

export function getStartupModeFlags(initialArgs: CliArgs | null | undefined): {
  shouldUseMinimalStartup: boolean;
  shouldSkipHeavyStartup: boolean;
} {
  return {
    shouldUseMinimalStartup: Boolean(
      (initialArgs && isStandaloneTexthookerCommand(initialArgs)) ||
      (initialArgs?.stats &&
        (initialArgs.statsCleanup || initialArgs.statsBackground || initialArgs.statsStop)),
    ),
    shouldSkipHeavyStartup: Boolean(
      initialArgs &&
      (shouldRunSettingsOnlyStartup(initialArgs) ||
        initialArgs.stats ||
        initialArgs.dictionary ||
        initialArgs.setup),
    ),
  };
}
