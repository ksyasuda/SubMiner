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
const MPV_WHEEL_KEYS = ['WHEEL_UP', 'WHEEL_DOWN', 'WHEEL_LEFT', 'WHEEL_RIGHT'] as const;
export type MpvWheelKey = (typeof MPV_WHEEL_KEYS)[number];
// DOM MouseEvent.button to mpv's mouse button names.
export const MPV_MOUSE_BUTTON_BY_BUTTON: Readonly<Record<number, string>> = {
  0: 'MBTN_LEFT',
  1: 'MBTN_MID',
  2: 'MBTN_RIGHT',
  3: 'MBTN_BACK',
  4: 'MBTN_FORWARD',
};
// mpv synthesizes these from two quick presses of the base button.
const MPV_DOUBLE_CLICK_KEYS = ['MBTN_LEFT_DBL', 'MBTN_MID_DBL', 'MBTN_RIGHT_DBL'];
const MPV_SPECIAL_KEYS = new Set<string>([
  ...Object.values(SPECIAL_KEYS),
  ...MPV_WHEEL_KEYS,
  ...Object.values(MPV_MOUSE_BUTTON_BY_BUTTON),
  ...MPV_DOUBLE_CLICK_KEYS,
]);
// Chromium reports 120 px per wheel notch on both Wayland and X11.
const WHEEL_NOTCH_PIXELS = 120;
const WHEEL_NOTCH_LINES = 3;
// Leading command flags accepted by mpv's input/cmd.c, before the command name.
const MPV_COMMAND_PREFIXES =
  /^(?:(?:no-osd|osd-bar|osd-msg|osd-msg-bar|osd-auto|expand-properties|raw|repeatable|nonrepeatable|nonscalable|async|sync)\s+)+/;

// Single keyboard strokes, mouse buttons, and wheel scrolls are imported. Key sequences
// and pointer motion (MOUSE_MOVE) are not.
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

// Maps a DOM wheel event to mpv's wheel key and its notch count, which mpv uses as the
// precise-scroll scale for `keypress <key> <scale>`. Trackpads yield fractional notches.
export function wheelEventToMpvWheel(
  event: Pick<WheelEvent, 'deltaX' | 'deltaY' | 'deltaMode'>,
): { key: MpvWheelKey; notches: number } | null {
  const vertical = Math.abs(event.deltaY) >= Math.abs(event.deltaX);
  const delta = vertical ? event.deltaY : event.deltaX;
  if (!Number.isFinite(delta) || delta === 0) return null;
  const key: MpvWheelKey = vertical
    ? delta < 0
      ? 'WHEEL_UP'
      : 'WHEEL_DOWN'
    : delta < 0
      ? 'WHEEL_LEFT'
      : 'WHEEL_RIGHT';
  // deltaMode: 0 = pixels, 1 = lines, 2 = pages.
  const unit =
    event.deltaMode === 1 ? WHEEL_NOTCH_LINES : event.deltaMode === 2 ? 1 : WHEEL_NOTCH_PIXELS;
  return { key, notches: Math.abs(delta) / unit };
}

export function parseMpvInputBindingKeys(
  value: unknown,
  { includeIgnored = true }: { includeIgnored?: boolean } = {},
): string[] {
  if (!Array.isArray(value)) return [];
  const bindings = new Map<string, { priority: number; owned: boolean; ignored: boolean }>();
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
      bindings.set(key, {
        priority: entry.priority,
        owned,
        ignored: entry.cmd.trim().replace(MPV_COMMAND_PREFIXES, '') === 'ignore',
      });
    }
  }
  return [...bindings]
    .filter(([, binding]) => !binding.owned && (includeIgnored || !binding.ignored))
    .map(([key]) => key);
}
