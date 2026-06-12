const PASSWORD_STORE_ARG = '--password-store';
const DEFAULT_LINUX_PASSWORD_STORE = 'gnome-libsecret';

export function getPasswordStoreArg(argv: string[]): string | null {
  let resolved: string | null = null;
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg?.startsWith(PASSWORD_STORE_ARG)) {
      continue;
    }

    if (arg === PASSWORD_STORE_ARG) {
      const value = argv[i + 1];
      if (value && !value.startsWith('--')) {
        resolved = value.trim();
        i += 1;
      }
      continue;
    }

    const [prefix, value] = arg.split('=', 2);
    if (prefix === PASSWORD_STORE_ARG && value && value.trim().length > 0) {
      resolved = value.trim();
    }
  }
  return resolved;
}

export function normalizePasswordStoreArg(value: string): string {
  const normalized = value.trim();
  if (normalized.toLowerCase() === 'gnome') {
    return DEFAULT_LINUX_PASSWORD_STORE;
  }
  return normalized;
}

export function getDefaultPasswordStore(): string {
  return DEFAULT_LINUX_PASSWORD_STORE;
}
