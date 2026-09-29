import assert from 'node:assert/strict';
import test from 'node:test';

import { dispatchConfiguredMpvCommand } from './mpv-command-dispatch';

async function dispatch(command: (string | number)[], paused: boolean | null | Error) {
  const sent: (string | number)[][] = [];
  dispatchConfiguredMpvCommand(command, {
    getPlaybackPaused: async () => {
      if (paused instanceof Error) throw paused;
      return paused;
    },
    sendMpvCommand: (mpvCommand) => sent.push(mpvCommand),
  });
  await new Promise((resolve) => setTimeout(resolve, 0));
  return sent;
}

test('subtitle seeks keep paused or unknown playback paused', async () => {
  const repaused = [
    ['sub-seek', 1],
    ['set_property', 'pause', 'yes'],
  ];
  assert.deepEqual(await dispatch(['sub-seek', 1], true), repaused);
  assert.deepEqual(await dispatch(['sub-seek', 1], null), repaused);
});

test('subtitle seeks leave running playback alone', async () => {
  assert.deepEqual(await dispatch(['sub-seek', -1], false), [['sub-seek', -1]]);
  assert.deepEqual(await dispatch(['sub-seek', -1], new Error('ipc down')), [['sub-seek', -1]]);
});

test('other mpv commands are sent as-is', async () => {
  assert.deepEqual(await dispatch(['cycle', 'pause'], true), [['cycle', 'pause']]);
});

test('a failed re-pause does not resend the subtitle seek', async () => {
  const sent: (string | number)[][] = [];
  const originalConsoleError = console.error;
  console.error = () => {};
  try {
    dispatchConfiguredMpvCommand(['sub-seek', 1], {
      getPlaybackPaused: async () => true,
      sendMpvCommand: (command) => {
        sent.push(command);
        if (command[0] === 'set_property') throw new Error('ipc closed');
      },
    });
    await new Promise((resolve) => setTimeout(resolve, 0));
  } finally {
    console.error = originalConsoleError;
  }

  assert.deepEqual(sent, [
    ['sub-seek', 1],
    ['set_property', 'pause', 'yes'],
  ]);
});
