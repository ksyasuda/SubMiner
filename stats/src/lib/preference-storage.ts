function getPreferenceStorage(): Storage | null {
  try {
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

export function readBooleanPreference(key: string, fallback: boolean): boolean {
  try {
    const value = getPreferenceStorage()?.getItem(key);
    if (value === 'true') return true;
    if (value === 'false') return false;
    return fallback;
  } catch {
    return fallback;
  }
}

export function writeBooleanPreference(key: string, value: boolean): void {
  try {
    getPreferenceStorage()?.setItem(key, String(value));
  } catch {
    // Storage can be blocked in private/restricted contexts; keep the in-memory choice.
  }
}
