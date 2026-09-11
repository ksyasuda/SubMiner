const SPECIAL_KEYS: Record<string, string> = {
  ' ': 'SPACE',
  '#': 'SHARP',
  Enter: 'ENTER',
  Escape: 'ESC',
  Backspace: 'BS',
  Tab: 'TAB',
  Delete: 'DEL',
  Insert: 'INS',
  Home: 'HOME',
  End: 'END',
  PageUp: 'PGUP',
  PageDown: 'PGDWN',
  ArrowLeft: 'LEFT',
  ArrowRight: 'RIGHT',
  ArrowUp: 'UP',
  ArrowDown: 'DOWN',
};
const MPV_SPECIAL_KEYS = new Set(Object.values(SPECIAL_KEYS));
// Leading command flags accepted by mpv's input/cmd.c, before the command name.
const MPV_COMMAND_PREFIXES =
  /^(?:(?:no-osd|osd-bar|osd-msg|osd-msg-bar|osd-auto|expand-properties|raw|repeatable|nonrepeatable|nonscalable|async|sync)\s+)+/;

// Only single keyboard strokes are imported. Mouse input and sequences need
// their own focus and conflict rules before they can be forwarded safely.
export function normalizeMpvInputKey(value: string): string | null {
  const modifiers = new Set<string>();
  let key = value;
  let modifier = /^(Shift|Ctrl|Alt|Meta)\+/i.exec(key);
  while (modifier?.[1]) {
    modifiers.add(modifier[1].toLowerCase());
    key = key.slice(modifier[0].length);
    modifier = /^(Shift|Ctrl|Alt|Meta)\+/i.exec(key);
  }
  if (key === 'SHARP') modifiers.delete('shift');
  if (!MPV_SPECIAL_KEYS.has(key) && !/^F(?:[1-9]|1[0-9]|2[0-4])$/.test(key)) {
    if ([...key].length !== 1) return null;
    if (modifiers.has('shift') && /^[a-z]$/i.test(key)) key = key.toUpperCase();
    modifiers.delete('shift');
  }
  return [...['ctrl', 'alt', 'shift', 'meta'].filter((item) => modifiers.has(item)), key].join('+');
}

export function keyboardEventToMpvKey(
  event: Pick<
    KeyboardEvent,
    'key' | 'ctrlKey' | 'altKey' | 'shiftKey' | 'metaKey' | 'isComposing' | 'getModifierState'
  >,
): string | null {
  if (event.isComposing || event.key === 'Dead' || event.getModifierState?.('AltGraph'))
    return null;
  const key = SPECIAL_KEYS[event.key] ?? event.key;
  const modifiers = [
    ...(event.ctrlKey ? ['ctrl'] : []),
    ...(event.altKey ? ['alt'] : []),
    ...(event.shiftKey ? ['shift'] : []),
    ...(event.metaKey ? ['meta'] : []),
  ];
  return normalizeMpvInputKey([...modifiers, key].join('+'));
}

export function parseMpvInputBindingKeys(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  const bindings = new Map<string, { priority: number; owned: boolean }>();
  for (const candidate of value) {
    const entry: unknown = candidate;
    if (
      !entry ||
      typeof entry !== 'object' ||
      !('key' in entry) ||
      typeof entry.key !== 'string' ||
      !('cmd' in entry) ||
      typeof entry.cmd !== 'string' ||
      !('priority' in entry) ||
      typeof entry.priority !== 'number' ||
      !Number.isFinite(entry.priority) ||
      entry.priority < 0
    )
      continue;
    const key = normalizeMpvInputKey(entry.key);
    if (!key) continue;
    const owner = 'owner' in entry ? entry.owner : undefined;
    const owned =
      owner === 'subminer' ||
      (owner === undefined &&
        /^(?:script-binding\s+["']?subminer\/|script-message\s+["']?subminer-)/.test(
          entry.cmd.trimStart().replace(MPV_COMMAND_PREFIXES, ''),
        ));
    const previous = bindings.get(key);
    // mpv's reported priority already ranks active non-weak bindings above weak
    // bindings. Only the winning binding determines whether the key is imported.
    if (
      !previous ||
      entry.priority > previous.priority ||
      (entry.priority === previous.priority && owned)
    ) {
      bindings.set(key, { priority: entry.priority, owned });
    }
  }
  return [...bindings].filter(([, binding]) => !binding.owned).map(([key]) => key);
}
