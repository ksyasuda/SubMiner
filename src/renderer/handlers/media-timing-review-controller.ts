import type { ResolvedControllerConfig } from '../../types';
import {
  createHoldState,
  normalizeRawButtonState,
  resetHeldAction,
  resolveDpadHorizontalValue,
  resolveDpadVerticalValue,
  shouldFireHeldAction,
  type ControllerScreenLayout,
  type GamepadLike,
  type HoldState,
} from './gamepad-controller.js';

/** Every button on the media timing review screen, named by what it does. */
export type MediaTimingReviewControllerAction =
  | 'confirm'
  | 'cancel'
  | 'play'
  | 'resetTiming'
  | 'showEarlier'
  | 'showLater'
  | 'startBack'
  | 'startForward'
  | 'endBack'
  | 'endForward'
  | 'previousLineAdd'
  | 'previousLineRemove'
  | 'nextLineAdd'
  | 'nextLineRemove'
  | 'framePrevious'
  | 'frameNext'
  | 'frameReset'
  | 'keepEditing'
  | 'useOriginal'
  | 'skipMedia'
  | 'discard';

/** The review shows either the timing editor or the "Stop reviewing?" step. */
export type MediaTimingReviewStep = 'editor' | 'cancel';

/** Bound by the letter printed on the button, so the layout matches what the user reads. */
type FaceButton = 'a' | 'b' | 'x' | 'y';
type FacePosition = 'south' | 'east' | 'west' | 'north';
type DpadDirection = 'left' | 'right' | 'up' | 'down';
type DpadModifier = 'none' | 'l1' | 'r1';

export type MediaTimingReviewControllerBinding =
  | { action: MediaTimingReviewControllerAction; step: MediaTimingReviewStep; face: FaceButton }
  | {
      action: MediaTimingReviewControllerAction;
      step: MediaTimingReviewStep;
      modifier: DpadModifier;
      direction: DpadDirection;
    }
  | { action: MediaTimingReviewControllerAction; step: MediaTimingReviewStep; shoulders: 'both' };

/** Where each letter sits: Xbox (and PlayStation) pads put A at the bottom, Nintendo-style pads on the right. */
const FACE_POSITIONS: Record<
  ResolvedControllerConfig['faceButtonLabels'],
  Record<FaceButton, FacePosition>
> = {
  xbox: { a: 'south', b: 'east', x: 'west', y: 'north' },
  nintendo: { a: 'east', b: 'south', x: 'north', y: 'west' },
};

/**
 * Fixed default layout. L1 holds everything about the clip start (and lines before it),
 * R1 everything about the clip end (and lines after it). D-pad entries repeat while held;
 * face buttons and the L1 + R1 chord fire once per press.
 */
export const MEDIA_TIMING_REVIEW_CONTROLLER_LAYOUT: readonly MediaTimingReviewControllerBinding[] =
  [
    { action: 'confirm', step: 'editor', face: 'a' },
    { action: 'cancel', step: 'editor', face: 'b' },
    { action: 'play', step: 'editor', face: 'x' },
    { action: 'resetTiming', step: 'editor', face: 'y' },
    { action: 'showEarlier', step: 'editor', modifier: 'none', direction: 'left' },
    { action: 'showLater', step: 'editor', modifier: 'none', direction: 'right' },
    { action: 'framePrevious', step: 'editor', modifier: 'none', direction: 'up' },
    { action: 'frameNext', step: 'editor', modifier: 'none', direction: 'down' },
    { action: 'frameReset', step: 'editor', shoulders: 'both' },
    { action: 'startBack', step: 'editor', modifier: 'l1', direction: 'left' },
    { action: 'startForward', step: 'editor', modifier: 'l1', direction: 'right' },
    { action: 'previousLineAdd', step: 'editor', modifier: 'l1', direction: 'up' },
    { action: 'previousLineRemove', step: 'editor', modifier: 'l1', direction: 'down' },
    { action: 'endBack', step: 'editor', modifier: 'r1', direction: 'left' },
    { action: 'endForward', step: 'editor', modifier: 'r1', direction: 'right' },
    { action: 'nextLineAdd', step: 'editor', modifier: 'r1', direction: 'up' },
    { action: 'nextLineRemove', step: 'editor', modifier: 'r1', direction: 'down' },
    { action: 'useOriginal', step: 'cancel', face: 'a' },
    { action: 'keepEditing', step: 'cancel', face: 'b' },
    { action: 'skipMedia', step: 'cancel', face: 'x' },
    { action: 'discard', step: 'cancel', face: 'y' },
  ];

const FACE_LABELS: Record<FaceButton, string> = { a: 'A', b: 'B', x: 'X', y: 'Y' };
const DIRECTION_LABELS: Record<DpadDirection, string> = {
  left: '◀',
  right: '▶',
  up: '▲',
  down: '▼',
};
const MODIFIER_LABELS: Record<DpadModifier, string> = { none: '', l1: 'L1 + ', r1: 'R1 + ' };

/** Short label shown next to a button, e.g. `A`, `◀`, `L1 + ▶`, `L1 + R1`. */
export function formatMediaTimingReviewControllerHint(
  binding: MediaTimingReviewControllerBinding,
): string {
  if ('face' in binding) return FACE_LABELS[binding.face];
  if ('shoulders' in binding) return 'L1 + R1';
  return `${MODIFIER_LABELS[binding.modifier]}${DIRECTION_LABELS[binding.direction]}`;
}

type ReviewInput = {
  face: Record<FaceButton, boolean>;
  l1: boolean;
  r1: boolean;
  horizontal: number;
  vertical: number;
};

type DpadChannel = {
  /** `modifier:direction` captured when the direction was pressed. */
  key: string | null;
  /** Set when the modifier changed mid-hold; the direction must be released first. */
  latched: boolean;
  hold: HoldState;
};

function createDpadChannel(): DpadChannel {
  return { key: null, latched: false, hold: createHoldState() };
}

function readButton(gamepad: GamepadLike, index: number, config: ResolvedControllerConfig) {
  return normalizeRawButtonState(gamepad.buttons[index], config.triggerDeadzone);
}

function readInput(gamepad: GamepadLike, config: ResolvedControllerConfig): ReviewInput {
  const indices = config.buttonIndices;
  const byPosition: Record<FacePosition, boolean> = {
    south: readButton(gamepad, indices.buttonSouth, config),
    east: readButton(gamepad, indices.buttonEast, config),
    west: readButton(gamepad, indices.buttonWest, config),
    north: readButton(gamepad, indices.buttonNorth, config),
  };
  const positions = FACE_POSITIONS[config.faceButtonLabels] ?? FACE_POSITIONS.xbox;
  return {
    face: {
      a: byPosition[positions.a],
      b: byPosition[positions.b],
      x: byPosition[positions.x],
      y: byPosition[positions.y],
    },
    l1: readButton(gamepad, indices.leftShoulder, config),
    r1: readButton(gamepad, indices.rightShoulder, config),
    horizontal: resolveDpadHorizontalValue(gamepad, config.triggerDeadzone),
    vertical: resolveDpadVerticalValue(gamepad, config.triggerDeadzone),
  };
}

function resolveModifier(input: ReviewInput): DpadModifier | 'both' {
  if (input.l1 && input.r1) return 'both';
  if (input.l1) return 'l1';
  if (input.r1) return 'r1';
  return 'none';
}

export function createMediaTimingReviewControllerLayout(options: {
  getStep: () => MediaTimingReviewStep;
  dispatch: (action: MediaTimingReviewControllerAction) => void;
}): ControllerScreenLayout {
  let previousFace: Record<FaceButton, boolean> = { a: false, b: false, x: false, y: false };
  let previousChord = false;
  const horizontal = createDpadChannel();
  const vertical = createDpadChannel();

  function findBinding(
    match: (binding: MediaTimingReviewControllerBinding) => boolean,
  ): MediaTimingReviewControllerAction | null {
    const step = options.getStep();
    const binding = MEDIA_TIMING_REVIEW_CONTROLLER_LAYOUT.find(
      (candidate) => candidate.step === step && match(candidate),
    );
    return binding?.action ?? null;
  }

  function fire(action: MediaTimingReviewControllerAction | null): void {
    if (action !== null) options.dispatch(action);
  }

  function updateDpadChannel(
    channel: DpadChannel,
    value: number,
    negative: DpadDirection,
    positive: DpadDirection,
    modifier: DpadModifier | 'both',
    config: ResolvedControllerConfig,
    now: number,
  ): void {
    if (value === 0) {
      channel.key = null;
      channel.latched = false;
      resetHeldAction(channel.hold);
      return;
    }
    if (channel.latched) return;

    const direction = value < 0 ? negative : positive;
    const key = `${modifier}:${direction}`;
    if (channel.key === null) {
      channel.key = key;
      resetHeldAction(channel.hold);
    } else if (channel.key !== key) {
      // Releasing L1 a moment before the D-pad must not fire the plain D-pad action.
      channel.latched = true;
      return;
    }

    if (!shouldFireHeldAction(channel.hold, now, config.repeatDelayMs, config.repeatIntervalMs)) {
      return;
    }
    if (modifier === 'both') return;
    fire(
      findBinding(
        (binding) =>
          'direction' in binding &&
          binding.modifier === modifier &&
          binding.direction === direction,
      ),
    );
  }

  function latch(input: ReviewInput): void {
    previousFace = { ...input.face };
    previousChord = input.l1 && input.r1;
    for (const [channel, value] of [
      [horizontal, input.horizontal],
      [vertical, input.vertical],
    ] as const) {
      channel.key = null;
      channel.latched = value !== 0;
      resetHeldAction(channel.hold);
    }
  }

  return {
    activate(gamepad, config) {
      latch(readInput(gamepad, config));
    },
    poll(gamepad, config, now) {
      const input = readInput(gamepad, config);

      for (const face of ['a', 'b', 'x', 'y'] as const) {
        if (input.face[face] && !previousFace[face]) {
          fire(findBinding((binding) => 'face' in binding && binding.face === face));
        }
      }
      previousFace = { ...input.face };

      const chord = input.l1 && input.r1;
      if (chord && !previousChord) {
        fire(findBinding((binding) => 'shoulders' in binding));
      }
      previousChord = chord;

      const modifier = resolveModifier(input);
      updateDpadChannel(horizontal, input.horizontal, 'left', 'right', modifier, config, now);
      updateDpadChannel(vertical, input.vertical, 'up', 'down', modifier, config, now);
    },
  };
}
