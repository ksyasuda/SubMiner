import type { GamepadLike } from './gamepad-controller.js';

/**
 * How far back a press the Gamepad API reported itself still counts as the same press the
 * main process is replaying. Covers the helper -> main -> renderer IPC hop.
 */
const SAME_PRESS_WINDOW_MS = 500;

/**
 * Replays macOS background presses (see `src/main/runtime/macos-controller-watch.ts`) into the
 * controller poll. A queued button reads as pressed for one poll, so the controller sees a
 * normal press edge; the next poll reads the real pad again, so a button still held does not
 * fire twice. A press the Gamepad API already reported is dropped.
 *
 * Only a lone standard-mapping pad gets the replay: the press carries no pad identity, so with
 * several visible it could run an action on the wrong one. With none visible yet, or a pad
 * mapped differently, the press only served to focus the overlay.
 */
export function createControllerPressReplay() {
  const pendingPresses: Array<{ buttonIndex: number; queuedAt: number }> = [];
  const lastRealPressAt = new Map<number, number>();

  function recordRealPresses(gamepads: GamepadLike[], now: number): void {
    for (const gamepad of gamepads) {
      gamepad.buttons.forEach((button, buttonIndex) => {
        if (button.pressed) lastRealPressAt.set(buttonIndex, now);
      });
    }
  }

  function withPressed(gamepad: GamepadLike, buttonIndices: number[]): GamepadLike {
    const buttons = gamepad.buttons.map((button, index) =>
      buttonIndices.includes(index) ? { value: 1, pressed: true, touched: true } : button,
    );
    return { ...gamepad, buttons };
  }

  return {
    queue(buttonIndex: number, now: number): void {
      pendingPresses.push({ buttonIndex, queuedAt: now });
    },

    /** Called once per poll with the Gamepad API's pads; returns the pads to read. */
    apply(gamepads: Array<GamepadLike | null>, now: number): Array<GamepadLike | null> {
      const connected = gamepads.filter((gamepad): gamepad is GamepadLike =>
        Boolean(gamepad?.connected),
      );
      recordRealPresses(connected, now);
      if (pendingPresses.length === 0) return gamepads;

      const replayed = pendingPresses
        .filter(({ buttonIndex, queuedAt }) => {
          const seenAt = lastRealPressAt.get(buttonIndex);
          return seenAt === undefined || seenAt < queuedAt - SAME_PRESS_WINDOW_MS;
        })
        .map(({ buttonIndex }) => buttonIndex);
      pendingPresses.length = 0;
      const onlyPad = connected.length === 1 ? connected[0] : undefined;
      if (replayed.length === 0 || onlyPad?.mapping !== 'standard') return gamepads;

      return gamepads.map((gamepad) =>
        gamepad === onlyPad ? withPressed(onlyPad, replayed) : gamepad,
      );
    },
  };
}
