import type { GamepadLike } from './gamepad-controller.js';

/**
 * How far apart the helper's report and the Gamepad API's report of one physical press can be.
 * The Gamepad API can lead (the helper -> main -> renderer IPC hop) or lag (Chromium only reads
 * the pad once the overlay has focus).
 */
const SAME_PRESS_WINDOW_MS = 500;

/**
 * Replays macOS background presses (see `src/main/runtime/macos-controller-watch.ts`) into the
 * controller poll. A queued button reads as pressed until the Gamepad API has reported that
 * press and its release, or until the window passes if it never does. Holding it across the gap
 * before a late report keeps one physical press to one press edge. A press the Gamepad API
 * already reported is dropped.
 *
 * Only a lone standard-mapping pad gets the replay: the press carries no pad identity, so with
 * several visible it could run an action on the wrong one. With none visible yet, or a pad
 * mapped differently, the press only served to focus the overlay.
 */
export function createControllerPressReplay() {
  const pendingPresses: Array<{ buttonIndex: number; queuedAt: number }> = [];
  const heldPresses = new Map<number, { releaseAt: number; realPressSeen: boolean }>();
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
    // Native Gamepad fields are prototype getters, so object spread would drop them.
    return {
      id: gamepad.id,
      index: gamepad.index,
      connected: gamepad.connected,
      mapping: gamepad.mapping,
      timestamp: gamepad.timestamp,
      axes: gamepad.axes,
      buttons,
    };
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
      if (pendingPresses.length === 0 && heldPresses.size === 0) return gamepads;

      for (const { buttonIndex, queuedAt } of pendingPresses) {
        const seenAt = lastRealPressAt.get(buttonIndex);
        if (seenAt === undefined || seenAt < queuedAt - SAME_PRESS_WINDOW_MS) {
          heldPresses.set(buttonIndex, {
            releaseAt: queuedAt + SAME_PRESS_WINDOW_MS,
            realPressSeen: false,
          });
        }
      }
      pendingPresses.length = 0;

      const onlyPad = connected.length === 1 ? connected[0] : undefined;
      if (onlyPad?.mapping !== 'standard') {
        heldPresses.clear();
        return gamepads;
      }

      for (const [buttonIndex, hold] of heldPresses) {
        if (onlyPad.buttons[buttonIndex]?.pressed) {
          hold.realPressSeen = true;
        } else if (hold.realPressSeen || now >= hold.releaseAt) {
          heldPresses.delete(buttonIndex);
        }
      }
      if (heldPresses.size === 0) return gamepads;

      const held = [...heldPresses.keys()];
      return gamepads.map((gamepad) =>
        gamepad === onlyPad ? withPressed(onlyPad, held) : gamepad,
      );
    },
  };
}
