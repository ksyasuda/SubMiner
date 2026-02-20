import { globalShortcut } from 'electron';

export function isGlobalShortcutRegisteredSafe(accelerator: string): boolean {
  try {
    return globalShortcut.isRegistered(accelerator);
  } catch {
    return false;
  }
}

export function shortcutMatchesInputForLocalFallback(
  input: Electron.Input,
  accelerator: string,
  allowWhenRegistered = false,
): boolean {
  if (input.type !== 'keyDown' || input.isAutoRepeat) return false;
  if (!accelerator) return false;
  if (!allowWhenRegistered && isGlobalShortcutRegisteredSafe(accelerator)) {
    return false;
  }

  const normalized = accelerator
    .replace(/\s+/g, '')
    .replace(/cmdorctrl/gi, 'CommandOrControl')
    .toLowerCase();
  const parts = normalized.split('+').filter(Boolean);
  if (parts.length === 0) return false;

  const keyToken = parts[parts.length - 1]!;
  const modifierTokens = new Set(parts.slice(0, -1));
  const allowedModifiers = new Set(['shift', 'alt', 'meta', 'control', 'commandorcontrol']);
  for (const token of modifierTokens) {
    if (!allowedModifiers.has(token)) return false;
  }

  const inputKey = (input.key || '').toLowerCase();
  if (keyToken.length === 1) {
    if (inputKey !== keyToken) return false;
  } else if (keyToken.startsWith('key') && keyToken.length === 4) {
    if (inputKey !== keyToken.slice(3)) return false;
  } else {
    return false;
  }

  const expectedShift = modifierTokens.has('shift');
  const expectedAlt = modifierTokens.has('alt');
  const expectedMeta = modifierTokens.has('meta');
  const expectedControl = modifierTokens.has('control');
  const expectedCommandOrControl = modifierTokens.has('commandorcontrol');

  if (Boolean(input.shift) !== expectedShift) return false;
  if (Boolean(input.alt) !== expectedAlt) return false;

  if (expectedCommandOrControl) {
    const hasCmdOrCtrl =
      process.platform === 'darwin' ? Boolean(input.meta || input.control) : Boolean(input.control);
    if (!hasCmdOrCtrl) return false;
  } else {
    if (process.platform === 'darwin') {
      if (input.meta || input.control) return false;
    } else if (!expectedControl && input.control) {
      return false;
    }
  }

  if (expectedMeta && !input.meta) return false;
  if (!expectedMeta && modifierTokens.has('meta') === false && input.meta) {
    if (!expectedCommandOrControl) return false;
  }

  if (expectedControl && !input.control) return false;
  if (!expectedControl && modifierTokens.has('control') === false && input.control) {
    if (!expectedCommandOrControl) return false;
  }

  return true;
}
