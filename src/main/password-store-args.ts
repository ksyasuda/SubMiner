export function getPasswordStoreArg(argv: string[]): string | null {
  let resolved: string | null = null;
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg?.startsWith('--password-store')) {
      continue;
    }

    if (arg === '--password-store') {
      const value = argv[i + 1];
      if (value && !value.startsWith('--')) {
        resolved = value.trim();
        i += 1;
      }
      continue;
    }

    const [prefix, value] = arg.split('=', 2);
    if (prefix === '--password-store' && value && value.trim().length > 0) {
      resolved = value.trim();
    }
  }
  return resolved;
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
