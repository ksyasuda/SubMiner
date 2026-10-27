import type {
  ControllerAxisDirection,
  ControllerButtonIndicesConfig,
  ControllerDpadDirection,
  ControllerFaceButtonLabels,
  ControllerFaceLetter,
  ControllerReviewAction,
  ControllerReviewInput,
  ResolvedControllerReviewBinding,
  ResolvedControllerReviewBindings,
} from '../types/runtime';

/**
 * A dialog with its own controller layout (media timing review, field grouping). Each
 * action presses one on-screen control on one step of the dialog; bindings are
 * press-or-combo, remappable in Alt+C, and stored under `controller.<configKey>`.
 */
export type ControllerScreenDefinition<A extends string = string, S extends string = string> = {
  id: 'review' | 'fieldGrouping';
  /** Alt+C group title. */
  title: string;
  /** Config key on `controller` and on each profile. */
  configKey: 'reviewBindings' | 'fieldGroupingBindings';
  /** Prefix for generated config descriptions, e.g. `Review dialog`. */
  describe: string;
  steps: readonly { id: S; label: string }[];
  actions: readonly ControllerScreenActionDefinition<A, S>[];
  defaults: Readonly<Record<A, ResolvedControllerReviewBinding>>;
};

export type ControllerScreenActionDefinition<
  A extends string = string,
  S extends string = string,
> = {
  id: A;
  label: string;
  step: S;
  /** Repeats while held, like a held arrow key. */
  repeat: boolean;
};

export type ControllerScreenBindings<A extends string = string> = Record<
  A,
  ResolvedControllerReviewBinding
>;

/** The review shows either the timing editor or the "Stop reviewing?" step. */
export type ControllerReviewStep = 'editor' | 'cancel';

export type ControllerReviewActionDefinition = ControllerScreenActionDefinition<
  ControllerReviewAction,
  ControllerReviewStep
>;

/** Every review action, in the order Alt+C lists them. */
export const CONTROLLER_REVIEW_ACTIONS: readonly ControllerReviewActionDefinition[] = [
  { id: 'confirm', label: 'Use this timing', step: 'editor', repeat: false },
  { id: 'cancel', label: 'Cancel', step: 'editor', repeat: false },
  { id: 'play', label: 'Play selection', step: 'editor', repeat: false },
  { id: 'resetTiming', label: 'Reset timing', step: 'editor', repeat: false },
  { id: 'showEarlier', label: 'Earlier −2s', step: 'editor', repeat: true },
  { id: 'showLater', label: 'Later +2s', step: 'editor', repeat: true },
  { id: 'framePrevious', label: 'Previous screenshot frame', step: 'editor', repeat: true },
  { id: 'frameNext', label: 'Next screenshot frame', step: 'editor', repeat: true },
  { id: 'frameReset', label: 'Reset screenshot frame', step: 'editor', repeat: false },
  { id: 'startBack', label: 'Start −0.1s', step: 'editor', repeat: true },
  { id: 'startForward', label: 'Start +0.1s', step: 'editor', repeat: true },
  { id: 'previousLineAdd', label: 'Add previous line', step: 'editor', repeat: true },
  { id: 'previousLineRemove', label: 'Remove previous line', step: 'editor', repeat: true },
  { id: 'endBack', label: 'End −0.1s', step: 'editor', repeat: true },
  { id: 'endForward', label: 'End +0.1s', step: 'editor', repeat: true },
  { id: 'nextLineAdd', label: 'Add next line', step: 'editor', repeat: true },
  { id: 'nextLineRemove', label: 'Remove next line', step: 'editor', repeat: true },
  { id: 'useOriginal', label: 'Use original timing', step: 'cancel', repeat: false },
  { id: 'keepEditing', label: 'Keep editing', step: 'cancel', repeat: false },
  { id: 'skipMedia', label: 'Without media', step: 'cancel', repeat: false },
  { id: 'discard', label: 'Discard card', step: 'cancel', repeat: false },
];

const ACTION_BY_ID = new Map(CONTROLLER_REVIEW_ACTIONS.map((action) => [action.id, action]));

export function getControllerReviewAction(
  id: ControllerReviewAction,
): ControllerReviewActionDefinition {
  return ACTION_BY_ID.get(id)!;
}

export function isControllerReviewAction(value: string): value is ControllerReviewAction {
  return ACTION_BY_ID.has(value as ControllerReviewAction);
}

export function getControllerScreenAction<A extends string, S extends string>(
  screen: ControllerScreenDefinition<A, S>,
  id: A,
): ControllerScreenActionDefinition<A, S> {
  return screen.actions.find((action) => action.id === id)!;
}

export function isControllerScreenAction<A extends string>(
  screen: ControllerScreenDefinition<A>,
  value: string,
): value is A {
  return screen.actions.some((action) => action.id === value);
}

export function getDefaultControllerScreenBindings<A extends string>(
  screen: ControllerScreenDefinition<A>,
): ControllerScreenBindings<A> {
  return structuredClone(screen.defaults) as ControllerScreenBindings<A>;
}

export const face = (letter: ControllerFaceLetter): ControllerReviewInput => ({
  kind: 'face',
  letter,
});
export const dpad = (direction: ControllerDpadDirection): ControllerReviewInput => ({
  kind: 'dpad',
  direction,
});
const LEFT_SHOULDER: ControllerReviewInput = { kind: 'button', buttonIndex: 4 };
const RIGHT_SHOULDER: ControllerReviewInput = { kind: 'button', buttonIndex: 5 };
export const combo = (
  press: ControllerReviewInput,
  hold: ControllerReviewInput | null = null,
): ResolvedControllerReviewBinding => ({ kind: 'combo', press, hold });

/**
 * Default layout. L1 holds everything about the clip start (and lines before it), R1
 * everything about the clip end (and lines after it). Face buttons follow the printed
 * letter and D-pad entries accept any D-pad report, so these work on every layout.
 */
const DEFAULT_CONTROLLER_REVIEW_BINDINGS: ResolvedControllerReviewBindings = {
  confirm: combo(face('a')),
  cancel: combo(face('b')),
  play: combo(face('x')),
  resetTiming: combo(face('y')),
  showEarlier: combo(dpad('left')),
  showLater: combo(dpad('right')),
  framePrevious: combo(dpad('up')),
  frameNext: combo(dpad('down')),
  frameReset: combo(RIGHT_SHOULDER, LEFT_SHOULDER),
  startBack: combo(dpad('left'), LEFT_SHOULDER),
  startForward: combo(dpad('right'), LEFT_SHOULDER),
  previousLineAdd: combo(dpad('up'), LEFT_SHOULDER),
  previousLineRemove: combo(dpad('down'), LEFT_SHOULDER),
  endBack: combo(dpad('left'), RIGHT_SHOULDER),
  endForward: combo(dpad('right'), RIGHT_SHOULDER),
  nextLineAdd: combo(dpad('up'), RIGHT_SHOULDER),
  nextLineRemove: combo(dpad('down'), RIGHT_SHOULDER),
  useOriginal: combo(face('a')),
  keepEditing: combo(face('b')),
  skipMedia: combo(face('x')),
  discard: combo(face('y')),
};

export const CONTROLLER_REVIEW_SCREEN: ControllerScreenDefinition<
  ControllerReviewAction,
  ControllerReviewStep
> = {
  id: 'review',
  title: 'Review media timing',
  configKey: 'reviewBindings',
  describe: 'Review dialog',
  steps: [
    { id: 'editor', label: 'While editing' },
    { id: 'cancel', label: 'After Cancel' },
  ],
  actions: CONTROLLER_REVIEW_ACTIONS,
  defaults: DEFAULT_CONTROLLER_REVIEW_BINDINGS,
};

export function getDefaultControllerReviewBindings(): ResolvedControllerReviewBindings {
  return structuredClone(DEFAULT_CONTROLLER_REVIEW_BINDINGS);
}

export function getDefaultControllerReviewBinding(
  action: ControllerReviewAction,
): ResolvedControllerReviewBinding {
  return structuredClone(DEFAULT_CONTROLLER_REVIEW_BINDINGS[action]);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isIndex(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0;
}

export function parseControllerReviewInput(value: unknown): ControllerReviewInput | null {
  if (!isRecord(value)) return null;
  switch (value.kind) {
    case 'button':
      return isIndex(value.buttonIndex) ? { kind: 'button', buttonIndex: value.buttonIndex } : null;
    case 'axis':
      return isIndex(value.axisIndex) &&
        (value.direction === 'negative' || value.direction === 'positive')
        ? { kind: 'axis', axisIndex: value.axisIndex, direction: value.direction }
        : null;
    case 'face':
      return value.letter === 'a' ||
        value.letter === 'b' ||
        value.letter === 'x' ||
        value.letter === 'y'
        ? { kind: 'face', letter: value.letter }
        : null;
    case 'dpad':
      return value.direction === 'left' ||
        value.direction === 'right' ||
        value.direction === 'up' ||
        value.direction === 'down'
        ? { kind: 'dpad', direction: value.direction }
        : null;
    default:
      return null;
  }
}

/** Accepts `{ "kind": "none" }` or `{ "press": …, "hold"?: … }` (`kind: "combo"` optional). */
export function parseControllerReviewBinding(
  value: unknown,
): ResolvedControllerReviewBinding | null {
  if (!isRecord(value)) return null;
  if (value.kind === 'none') return { kind: 'none' };
  if (value.kind !== undefined && value.kind !== 'combo') return null;
  const press = parseControllerReviewInput(value.press);
  if (!press) return null;
  if (value.hold === undefined || value.hold === null) return combo(press);
  const hold = parseControllerReviewInput(value.hold);
  return hold ? combo(press, hold) : null;
}

/** What a review input needs to know about the controller it runs on. */
export type ControllerReviewInputContext = {
  /** Gamepad API `mapping` ("standard" or ""); unknown before the pad is seen. */
  mapping: string | null | undefined;
  faceButtonLabels: ControllerFaceButtonLabels;
  buttonIndices: Required<ControllerButtonIndicesConfig>;
};

/** A physical input after resolving letters and D-pad reports; equal sources are the same input. */
export type ControllerReviewInputSource =
  | { kind: 'button'; buttonIndex: number }
  | { kind: 'axis'; axisIndex: number; direction: ControllerAxisDirection }
  | { kind: 'dpad'; direction: ControllerDpadDirection };

type FacePosition = 'south' | 'east' | 'west' | 'north';

/**
 * Where each layout letter sits. Xbox pads put A at the bottom, Nintendo-style pads on the
 * right. PlayStation pads follow the Xbox layout, so A is ×, B ○, X □ and Y △.
 */
const FACE_POSITIONS: Record<
  ControllerFaceButtonLabels,
  Record<ControllerFaceLetter, FacePosition>
> = {
  xbox: { a: 'south', b: 'east', x: 'west', y: 'north' },
  nintendo: { a: 'east', b: 'south', x: 'north', y: 'west' },
  playstation: { a: 'south', b: 'east', x: 'west', y: 'north' },
};

/** Every face button style, in the order Alt+C offers them. */
export const CONTROLLER_FACE_BUTTON_STYLES: readonly {
  id: ControllerFaceButtonLabels;
  name: string;
  /** Tells the layouts apart, since 8BitDo and other pads come in more than one. */
  hint: string;
}[] = [
  { id: 'xbox', name: 'Xbox', hint: 'A at the bottom, B on the right' },
  { id: 'nintendo', name: 'Nintendo', hint: 'B at the bottom, A on the right' },
  { id: 'playstation', name: 'PlayStation', hint: '\u00d7 at the bottom, \u25cb on the right' },
];

export function isControllerFaceButtonLabels(value: unknown): value is ControllerFaceButtonLabels {
  return CONTROLLER_FACE_BUTTON_STYLES.some((style) => style.id === value);
}

const POSITION_INDEX_KEYS = {
  south: 'buttonSouth',
  east: 'buttonEast',
  west: 'buttonWest',
  north: 'buttonNorth',
} as const satisfies Record<FacePosition, keyof ControllerButtonIndicesConfig>;

/** Unknown (pad not seen yet) counts as standard, the layout most pads report. */
function isStandardMapping(mapping: string | null | undefined): boolean {
  return mapping === 'standard' || mapping === null || mapping === undefined;
}

const STANDARD_DPAD_BUTTONS: Record<number, ControllerDpadDirection> = {
  12: 'up',
  13: 'down',
  14: 'left',
  15: 'right',
};

export function resolveControllerReviewInput(
  input: ControllerReviewInput,
  context: ControllerReviewInputContext,
): ControllerReviewInputSource {
  if (input.kind === 'dpad') return input;
  if (input.kind === 'face') {
    const positions = FACE_POSITIONS[context.faceButtonLabels] ?? FACE_POSITIONS.xbox;
    const buttonIndex = context.buttonIndices[POSITION_INDEX_KEYS[positions[input.letter]]];
    return { kind: 'button', buttonIndex };
  }
  if (input.kind === 'button') {
    const direction = isStandardMapping(context.mapping)
      ? STANDARD_DPAD_BUTTONS[input.buttonIndex]
      : undefined;
    return direction ? { kind: 'dpad', direction } : input;
  }
  // Non-standard pads usually report the D-pad as axes 6 (left/right) and 7 (up/down).
  if (!isStandardMapping(context.mapping) && (input.axisIndex === 6 || input.axisIndex === 7)) {
    const negative = input.axisIndex === 6 ? 'left' : 'up';
    const positive = input.axisIndex === 6 ? 'right' : 'down';
    return { kind: 'dpad', direction: input.direction === 'negative' ? negative : positive };
  }
  return input;
}

export function controllerReviewInputSourceKey(source: ControllerReviewInputSource): string {
  if (source.kind === 'dpad') return `dpad:${source.direction}`;
  if (source.kind === 'button') return `button:${source.buttonIndex}`;
  return `axis:${source.axisIndex}:${source.direction}`;
}

export type CompiledControllerScreenBinding<A extends string = string> = {
  action: A;
  repeat: boolean;
  press: ControllerReviewInputSource;
  pressKey: string;
  hold: ControllerReviewInputSource | null;
  holdKey: string | null;
};

export type CompiledControllerReviewBinding =
  CompiledControllerScreenBinding<ControllerReviewAction>;

/** The bindings active on one step of a screen, resolved to physical inputs, in Alt+C order. */
export function compileControllerScreenBindings<A extends string, S extends string>(
  screen: ControllerScreenDefinition<A, S>,
  bindings: Partial<ControllerScreenBindings<A>>,
  step: S,
  context: ControllerReviewInputContext,
): CompiledControllerScreenBinding<A>[] {
  const compiled: CompiledControllerScreenBinding<A>[] = [];
  for (const action of screen.actions) {
    if (action.step !== step) continue;
    const binding = bindings[action.id];
    if (!binding || binding.kind !== 'combo') continue;
    const press = resolveControllerReviewInput(binding.press, context);
    const hold = binding.hold ? resolveControllerReviewInput(binding.hold, context) : null;
    compiled.push({
      action: action.id,
      repeat: action.repeat,
      press,
      pressKey: controllerReviewInputSourceKey(press),
      hold,
      holdKey: hold ? controllerReviewInputSourceKey(hold) : null,
    });
  }
  return compiled;
}

/** The review bindings active on one step, resolved to physical inputs, in Alt+C order. */
export function compileControllerReviewBindings(
  bindings: ResolvedControllerReviewBindings,
  step: ControllerReviewStep,
  context: ControllerReviewInputContext,
): CompiledControllerReviewBinding[] {
  return compileControllerScreenBindings(CONTROLLER_REVIEW_SCREEN, bindings, step, context);
}

const DPAD_GLYPHS: Record<ControllerDpadDirection, string> = {
  left: '◀',
  right: '▶',
  up: '▲',
  down: '▼',
};

const FACE_POSITION_BY_INDEX: Record<number, FacePosition> = {
  0: 'south',
  1: 'east',
  2: 'west',
  3: 'north',
};

/** What is printed at each position, per style. */
const LETTER_AT_POSITION: Record<ControllerFaceButtonLabels, Record<FacePosition, string>> = {
  xbox: { south: 'A', east: 'B', west: 'X', north: 'Y' },
  nintendo: { south: 'B', east: 'A', west: 'Y', north: 'X' },
  playstation: { south: '\u00d7', east: '\u25cb', west: '\u25a1', north: '\u25b3' },
};

/** The printed name of the face button at index 0-3 (south, east, west, north), if any. */
export function getControllerFaceButtonName(
  buttonIndex: number,
  faceButtonLabels: ControllerFaceButtonLabels,
): string | null {
  const position = FACE_POSITION_BY_INDEX[buttonIndex];
  if (!position) return null;
  return (LETTER_AT_POSITION[faceButtonLabels] ?? LETTER_AT_POSITION.xbox)[position];
}

/** How a layout letter is printed on this style, e.g. `a` is `A` on Xbox and `×` on PlayStation. */
function formatFaceLetter(
  letter: ControllerFaceLetter,
  faceButtonLabels: ControllerFaceButtonLabels,
): string {
  const positions = FACE_POSITIONS[faceButtonLabels] ?? FACE_POSITIONS.xbox;
  return (LETTER_AT_POSITION[faceButtonLabels] ?? LETTER_AT_POSITION.xbox)[positions[letter]];
}

const STANDARD_BUTTON_NAMES: Record<number, string> = {
  4: 'L1',
  5: 'R1',
  6: 'L2',
  7: 'R2',
  8: 'Select',
  9: 'Start',
  10: 'L3',
  11: 'R3',
  16: 'Home',
};

const RAW_BUTTON_NAMES: Record<number, string> = {
  4: 'L1',
  5: 'R1',
  6: 'Back',
  7: 'Start',
  8: 'L3',
  9: 'R3',
};

type AxisName = { name: string; horizontal: boolean } | { name: string; trigger: true };

const STANDARD_AXIS_NAMES: Record<number, AxisName> = {
  0: { name: 'LS', horizontal: true },
  1: { name: 'LS', horizontal: false },
  2: { name: 'RS', horizontal: true },
  3: { name: 'RS', horizontal: false },
};

const RAW_AXIS_NAMES: Record<number, AxisName> = {
  0: { name: 'LS', horizontal: true },
  1: { name: 'LS', horizontal: false },
  2: { name: 'L2', trigger: true },
  3: { name: 'RS', horizontal: true },
  4: { name: 'RS', horizontal: false },
  5: { name: 'R2', trigger: true },
};

function formatSource(
  source: ControllerReviewInputSource,
  context: ControllerReviewInputContext,
  dpadPrefix: boolean,
): string {
  if (source.kind === 'dpad') {
    return `${dpadPrefix ? 'D-pad ' : ''}${DPAD_GLYPHS[source.direction]}`;
  }
  const standard = isStandardMapping(context.mapping);
  if (source.kind === 'button') {
    const faceName = getControllerFaceButtonName(source.buttonIndex, context.faceButtonLabels);
    if (faceName) return faceName;
    const names = standard ? STANDARD_BUTTON_NAMES : RAW_BUTTON_NAMES;
    return names[source.buttonIndex] ?? `Button ${source.buttonIndex}`;
  }
  const axis = (standard ? STANDARD_AXIS_NAMES : RAW_AXIS_NAMES)[source.axisIndex];
  const sign = source.direction === 'negative' ? '−' : '+';
  if (!axis) return `Axis ${source.axisIndex} ${sign}`;
  if ('trigger' in axis) return axis.name;
  const glyph = axis.horizontal
    ? source.direction === 'negative'
      ? DPAD_GLYPHS.left
      : DPAD_GLYPHS.right
    : source.direction === 'negative'
      ? DPAD_GLYPHS.up
      : DPAD_GLYPHS.down;
  return `${axis.name} ${glyph}`;
}

/** Short name for one input, e.g. `A`, `L1`, `◀` (or `D-pad ◀`), `LS ▲`. */
export function formatControllerReviewInput(
  input: ControllerReviewInput,
  context: ControllerReviewInputContext,
  options: { dpadPrefix?: boolean } = {},
): string {
  // Letters are shown as printed, whatever position they resolve to.
  if (input.kind === 'face') return formatFaceLetter(input.letter, context.faceButtonLabels);
  return formatSource(
    resolveControllerReviewInput(input, context),
    context,
    options.dpadPrefix ?? false,
  );
}

/** `None`, `A`, or `L1 + D-pad ▲` style label for a whole binding. */
export function formatControllerReviewBinding(
  binding: ResolvedControllerReviewBinding,
  context: ControllerReviewInputContext,
  options: { dpadPrefix?: boolean } = {},
): string {
  if (binding.kind === 'none') return 'None';
  const press = formatControllerReviewInput(binding.press, context, options);
  return binding.hold
    ? `${formatControllerReviewInput(binding.hold, context, options)} + ${press}`
    : press;
}

/**
 * Problems Alt+C should point out, per action: two actions on the same input in one step,
 * or a button used alone that other actions on that step use as their hold button.
 */
export function findControllerScreenBindingIssues<A extends string, S extends string>(
  screen: ControllerScreenDefinition<A, S>,
  bindings: Partial<ControllerScreenBindings<A>>,
  context: ControllerReviewInputContext,
): Partial<Record<A, string>> {
  const issues: Partial<Record<A, string>> = {};
  const label = (id: A) => getControllerScreenAction(screen, id).label;
  for (const step of screen.steps) {
    const compiled = compileControllerScreenBindings(screen, bindings, step.id, context);
    const holdKeys = new Map<string, ControllerReviewInputSource>();
    for (const entry of compiled) {
      if (entry.hold && entry.holdKey) holdKeys.set(entry.holdKey, entry.hold);
    }
    const firstByInput = new Map<string, A>();
    for (const entry of compiled) {
      if (entry.holdKey === entry.pressKey) {
        issues[entry.action] = 'Hold and press are the same button';
        continue;
      }
      const inputKey = `${entry.holdKey ?? ''}|${entry.pressKey}`;
      const first = firstByInput.get(inputKey);
      if (first) {
        issues[entry.action] = `Same as ${label(first)}`;
        issues[first] ??= `Same as ${label(entry.action)}`;
        continue;
      }
      firstByInput.set(inputKey, entry.action);
      const holdSource = entry.holdKey === null ? holdKeys.get(entry.pressKey) : undefined;
      if (holdSource) {
        issues[entry.action] =
          `${formatSource(holdSource, context, true)} is also a hold button on this step`;
      }
    }
  }
  return issues;
}

export function findControllerReviewBindingIssues(
  bindings: ResolvedControllerReviewBindings,
  context: ControllerReviewInputContext,
): Partial<Record<ControllerReviewAction, string>> {
  return findControllerScreenBindingIssues(CONTROLLER_REVIEW_SCREEN, bindings, context);
}
