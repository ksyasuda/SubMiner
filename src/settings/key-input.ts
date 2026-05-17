import type { Keybinding } from '../types/runtime';

export type KeyInputMode = 'accelerator' | 'dom-code' | 'code';

export interface KeyboardInputLike {
  code: string;
  key: string;
  ctrlKey: boolean;
  altKey: boolean;
  shiftKey: boolean;
  metaKey: boolean;
}

export interface MpvKeybindingRow {
  defaultKey: string;
  key: string;
  command: (string | number)[] | null;
  commandText: string;
  isDefault: boolean;
}

const MODIFIER_CODES = new Set([
  'AltLeft',
  'AltRight',
  'ControlLeft',
  'ControlRight',
  'MetaLeft',
  'MetaRight',
  'ShiftLeft',
  'ShiftRight',
]);

const ELECTRON_KEY_BY_CODE: Record<string, string> = {
  Backquote: 'Backquote',
  Backslash: 'Backslash',
  BracketLeft: 'BracketLeft',
  BracketRight: 'BracketRight',
  Comma: 'Comma',
  Delete: 'Delete',
  End: 'End',
  Enter: 'Enter',
  Equal: 'Plus',
  Escape: 'Escape',
  Home: 'Home',
  Insert: 'Insert',
  Minus: 'Minus',
  PageDown: 'PageDown',
  PageUp: 'PageUp',
  Period: 'Period',
  Quote: 'Quote',
  Semicolon: 'Semicolon',
  Slash: 'Slash',
  Space: 'Space',
  Tab: 'Tab',
};

function commandEquals(a: Keybinding['command'], b: Keybinding['command']): boolean {
  return JSON.stringify(a) === JSON.stringify(b);
}

function normalizeUserBindings(userBindings: unknown): Keybinding[] {
  if (!Array.isArray(userBindings)) return [];
  return userBindings.filter((binding): binding is Keybinding => {
    if (!binding || typeof binding !== 'object') return false;
    const candidate = binding as Keybinding;
    return (
      typeof candidate.key === 'string' &&
      (candidate.command === null || Array.isArray(candidate.command))
    );
  });
}

function electronKeyToken(input: KeyboardInputLike): string | null {
  if (/^Key[A-Z]$/.test(input.code)) return input.code.slice(3);
  if (/^Digit[0-9]$/.test(input.code)) return input.code.slice(5);
  if (/^Numpad[0-9]$/.test(input.code)) return `num${input.code.slice(6)}`;
  if (/^F\d{1,2}$/.test(input.code)) return input.code;
  if (input.code.startsWith('Arrow')) return input.code.replace('Arrow', '');
  return ELECTRON_KEY_BY_CODE[input.code] ?? null;
}

export function keyboardEventToConfigKey(
  input: KeyboardInputLike,
  mode: KeyInputMode,
): string | null {
  if (!input.code || MODIFIER_CODES.has(input.code)) {
    return null;
  }

  if (mode === 'code') {
    return input.code;
  }

  const parts: string[] = [];
  if (mode === 'accelerator') {
    if (input.ctrlKey || input.metaKey) parts.push('CommandOrControl');
    if (input.altKey) parts.push('Alt');
    if (input.shiftKey) parts.push('Shift');
    const key = electronKeyToken(input);
    return key ? [...parts, key].join('+') : null;
  }

  if (input.ctrlKey) parts.push('Ctrl');
  if (input.altKey) parts.push('Alt');
  if (input.shiftKey) parts.push('Shift');
  if (input.metaKey) parts.push('Meta');
  return [...parts, input.code].join('+');
}

export function createMpvKeybindingRows(
  defaultBindings: Keybinding[],
  userBindings: unknown,
): MpvKeybindingRow[] {
  const normalizedUserBindings = normalizeUserBindings(userBindings);
  const userByKey = new Map(normalizedUserBindings.map((binding) => [binding.key, binding]));
  const consumedUserKeys = new Set<string>();

  const rows = defaultBindings.map((binding) => {
    const override = userByKey.get(binding.key);
    if (override?.command === null) {
      const movedOverride = normalizedUserBindings.find(
        (candidate) =>
          candidate.key !== binding.key && commandEquals(candidate.command, binding.command),
      );
      if (movedOverride) {
        consumedUserKeys.add(binding.key);
        consumedUserKeys.add(movedOverride.key);
        return {
          defaultKey: binding.key,
          key: movedOverride.key,
          command: movedOverride.command,
          commandText: JSON.stringify(movedOverride.command ?? null),
          isDefault: true,
        };
      }
    }

    if (override) {
      consumedUserKeys.add(binding.key);
    }
    const command = override?.command ?? binding.command;
    return {
      defaultKey: binding.key,
      key: binding.key,
      command,
      commandText: JSON.stringify(command ?? null),
      isDefault: true,
    };
  });

  for (const binding of normalizedUserBindings) {
    if (consumedUserKeys.has(binding.key)) {
      continue;
    }
    if (defaultBindings.some((defaultBinding) => defaultBinding.key === binding.key)) {
      continue;
    }
    rows.push({
      defaultKey: binding.key,
      key: binding.key,
      command: binding.command,
      commandText: JSON.stringify(binding.command ?? null),
      isDefault: false,
    });
  }

  return rows;
}

export function parseMpvCommandText(value: string): Keybinding['command'] | undefined {
  try {
    const parsed = JSON.parse(value);
    if (parsed === null) return null;
    if (
      Array.isArray(parsed) &&
      parsed.every((entry) => typeof entry === 'string' || typeof entry === 'number')
    ) {
      return parsed;
    }
  } catch {
    return undefined;
  }
  return undefined;
}

export function buildMpvKeybindingConfigValue(
  defaultBindings: Keybinding[],
  rows: MpvKeybindingRow[],
): Keybinding[] {
  const next: Keybinding[] = [];

  for (const defaultBinding of defaultBindings) {
    const row = rows.find((candidate) => candidate.defaultKey === defaultBinding.key);
    if (!row) {
      next.push({ key: defaultBinding.key, command: null });
      continue;
    }

    if (row.key !== defaultBinding.key) {
      next.push({ key: defaultBinding.key, command: null });
      if (row.command !== null) {
        next.push({ key: row.key, command: row.command });
      }
      continue;
    }

    if (!commandEquals(row.command, defaultBinding.command)) {
      next.push({ key: row.key, command: row.command });
    }
  }

  for (const row of rows.filter((candidate) => !candidate.isDefault)) {
    next.push({ key: row.key, command: row.command });
  }

  return next;
}
