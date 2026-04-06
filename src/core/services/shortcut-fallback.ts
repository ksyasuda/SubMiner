import electron from 'electron';

const { globalShortcut } = electron;

export function isGlobalShortcutRegisteredSafe(accelerator: string): boolean {
  try {
    return globalShortcut.isRegistered(accelerator);
  } catch {
    return false;
  }
}

function matchesAcceleratorKeyToken(input: Electron.Input, keyToken: string): boolean {
  const inputCode = typeof input.code === 'string' ? input.code.toLowerCase() : '';
  const inputKey = typeof input.key === 'string' ? input.key.toLowerCase() : '';

  if (keyToken.length === 1) {
    if (/^[a-z]$/.test(keyToken)) {
      return inputCode === `key${keyToken}` || inputKey === keyToken;
    }
    if (/^[0-9]$/.test(keyToken)) {
      return inputCode === `digit${keyToken}` || inputKey === keyToken;
    }
    return inputKey === keyToken;
  }

  if (keyToken.startsWith('key') && keyToken.length === 4) {
    return inputCode === keyToken || inputKey === keyToken.slice(3);
  }
  if (keyToken.startsWith('digit') && keyToken.length === 6) {
    return inputCode === keyToken || inputKey === keyToken.slice(5);
  }
  if (/^f\d{1,2}$/.test(keyToken)) {
    return inputCode === keyToken || inputKey === keyToken;
  }

  const mappedTokens: Record<string, { codes: string[]; keys: string[] }> = {
    space: { codes: ['space'], keys: [' '] },
    tab: { codes: ['tab'], keys: ['tab'] },
    enter: { codes: ['enter'], keys: ['enter'] },
    return: { codes: ['enter'], keys: ['enter'] },
    esc: { codes: ['escape'], keys: ['escape'] },
    escape: { codes: ['escape'], keys: ['escape'] },
    up: { codes: ['arrowup'], keys: ['arrowup'] },
    down: { codes: ['arrowdown'], keys: ['arrowdown'] },
    left: { codes: ['arrowleft'], keys: ['arrowleft'] },
    right: { codes: ['arrowright'], keys: ['arrowright'] },
    backspace: { codes: ['backspace'], keys: ['backspace'] },
    delete: { codes: ['delete'], keys: ['delete'] },
    slash: { codes: ['slash'], keys: ['/'] },
    backslash: { codes: ['backslash'], keys: ['\\'] },
    minus: { codes: ['minus'], keys: ['-'] },
    plus: { codes: ['equal'], keys: ['+'] },
    equal: { codes: ['equal'], keys: ['='] },
    comma: { codes: ['comma'], keys: [','] },
    period: { codes: ['period'], keys: ['.'] },
    quote: { codes: ['quote'], keys: ["'"] },
    semicolon: { codes: ['semicolon'], keys: [';'] },
    bracketleft: { codes: ['bracketleft'], keys: ['['] },
    bracketright: { codes: ['bracketright'], keys: [']'] },
    backquote: { codes: ['backquote'], keys: ['`'] },
  };

  const mapping = mappedTokens[keyToken];
  if (!mapping) {
    return false;
  }

  return mapping.codes.includes(inputCode) || mapping.keys.includes(inputKey);
}

function normalizeModifierToken(token: string): string {
  if (token === 'ctrl') {
    return 'control';
  }
  if (token === 'option') {
    return 'alt';
  }
  if (token === 'cmd' || token === 'command' || token === 'super') {
    return 'meta';
  }
  return token;
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
  const modifierTokens = new Set(parts.slice(0, -1).map(normalizeModifierToken));
  const allowedModifiers = new Set(['shift', 'alt', 'meta', 'control', 'commandorcontrol']);
  for (const token of modifierTokens) {
    if (!allowedModifiers.has(token)) return false;
  }

  if (!matchesAcceleratorKeyToken(input, keyToken)) {
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
