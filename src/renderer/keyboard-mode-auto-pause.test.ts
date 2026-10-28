import assert from 'node:assert/strict';
import test from 'node:test';

import { createKeyboardModeAutoPause } from './keyboard-mode-auto-pause';

function setup(options: { enabled?: boolean; paused?: boolean | null } = {}) {
  const commands: string[] = [];
  const state = {
    enabled: options.enabled ?? true,
    paused: options.paused === undefined ? false : options.paused,
  };
  const autoPause = createKeyboardModeAutoPause({
    isEnabled: () => state.enabled,
    getPlaybackPaused: async () => state.paused,
    sendMpvCommand: (command) => {
      commands.push(command.join(' '));
      if (command[1] === 'pause') {
        state.paused = command[2] === 'yes';
        // mpv reports the new pause state back.
        autoPause.handlePlaybackPauseChanged(state.paused);
      }
    },
  });
  /** The user pauses or plays from mpv, the controller or the keyboard. */
  const userSetsPause = (paused: boolean) => {
    state.paused = paused;
    autoPause.handlePlaybackPauseChanged(paused);
  };
  return { autoPause, commands, state, userSetsPause };
}

test('entering keyboard-only mode pauses a playing video and leaving resumes it', async () => {
  const { autoPause, commands } = setup();

  await autoPause.handleModeChanged(true);
  assert.equal(autoPause.isPausedByMode(), true);
  await autoPause.handleModeChanged(false);

  assert.deepEqual(commands, ['set_property pause yes', 'set_property pause no']);
});

test('the setting is off: keyboard-only mode leaves playback alone', async () => {
  const { autoPause, commands } = setup({ enabled: false });

  await autoPause.handleModeChanged(true);
  await autoPause.handleModeChanged(false);

  assert.deepEqual(commands, []);
});

test('an already paused video stays paused when leaving the mode', async () => {
  const { autoPause, commands } = setup({ paused: true });

  await autoPause.handleModeChanged(true);
  await autoPause.handleModeChanged(false);

  assert.deepEqual(commands, []);
});

test('playing the video yourself in the mode means leaving does not touch it', async () => {
  const { autoPause, commands, userSetsPause } = setup();

  await autoPause.handleModeChanged(true);
  userSetsPause(false);
  await autoPause.handleModeChanged(false);

  assert.deepEqual(commands, ['set_property pause yes']);
});

test('playing and pausing again yourself also keeps the video paused on leaving', async () => {
  const { autoPause, commands, state, userSetsPause } = setup();

  await autoPause.handleModeChanged(true);
  userSetsPause(false);
  userSetsPause(true);
  await autoPause.handleModeChanged(false);

  assert.deepEqual(commands, ['set_property pause yes']);
  assert.equal(state.paused, true);
});

test('leaving before the pause lookup returns does not pause afterwards', async () => {
  const commands: string[] = [];
  let answer: (paused: boolean) => void = () => {};
  const autoPause = createKeyboardModeAutoPause({
    isEnabled: () => true,
    getPlaybackPaused: () => new Promise((resolve) => (answer = resolve)),
    sendMpvCommand: (command) => commands.push(command.join(' ')),
  });

  const entering = autoPause.handleModeChanged(true);
  await autoPause.handleModeChanged(false);
  answer(false);
  await entering;

  assert.deepEqual(commands, []);
  assert.equal(autoPause.isPausedByMode(), false);
});

test('an unknown pause state is left alone', async () => {
  const { autoPause, commands } = setup({ paused: null });

  await autoPause.handleModeChanged(true);
  await autoPause.handleModeChanged(false);

  assert.deepEqual(commands, []);
});

test('leaving the mode with the dictionary popup open leaves the pause to the popup', async () => {
  const commands: string[] = [];
  let popupOwnsPause = false;
  const autoPause = createKeyboardModeAutoPause({
    isEnabled: () => true,
    getPlaybackPaused: async () => false,
    sendMpvCommand: (command) => commands.push(command.join(' ')),
    handOffPause: () => {
      popupOwnsPause = true;
      return true;
    },
  });

  await autoPause.handleModeChanged(true);
  autoPause.handlePlaybackPauseChanged(true);
  await autoPause.handleModeChanged(false);

  assert.deepEqual(commands, ['set_property pause yes']);
  assert.equal(popupOwnsPause, true);
});
