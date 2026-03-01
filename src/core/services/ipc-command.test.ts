import assert from 'node:assert/strict';
import test from 'node:test';
import { handleMpvCommandFromIpc } from './ipc-command';

function createOptions(overrides: Partial<Parameters<typeof handleMpvCommandFromIpc>[1]> = {}) {
  const calls: string[] = [];
  const sentCommands: (string | number)[][] = [];
  const osd: string[] = [];
  const options: Parameters<typeof handleMpvCommandFromIpc>[1] = {
    specialCommands: {
      SUBSYNC_TRIGGER: '__subsync-trigger',
      RUNTIME_OPTIONS_OPEN: '__runtime-options-open',
      RUNTIME_OPTION_CYCLE_PREFIX: '__runtime-option-cycle:',
      REPLAY_SUBTITLE: '__replay-subtitle',
      PLAY_NEXT_SUBTITLE: '__play-next-subtitle',
    },
    triggerSubsyncFromConfig: () => {
      calls.push('subsync');
    },
    openRuntimeOptionsPalette: () => {
      calls.push('runtime-options');
    },
    runtimeOptionsCycle: () => ({ ok: true }),
    showMpvOsd: (text) => {
      osd.push(text);
    },
    mpvReplaySubtitle: () => {
      calls.push('replay');
    },
    mpvPlayNextSubtitle: () => {
      calls.push('next');
    },
    mpvSendCommand: (command) => {
      sentCommands.push(command);
    },
    isMpvConnected: () => true,
    hasRuntimeOptionsManager: () => true,
    ...overrides,
  };
  return { options, calls, sentCommands, osd };
}

test('handleMpvCommandFromIpc forwards regular mpv commands', () => {
  const { options, sentCommands, osd } = createOptions();
  handleMpvCommandFromIpc(['cycle', 'pause'], options);
  assert.deepEqual(sentCommands, [['cycle', 'pause']]);
  assert.deepEqual(osd, []);
});

test('handleMpvCommandFromIpc emits osd for subtitle position keybinding proxies', () => {
  const { options, sentCommands, osd } = createOptions();
  handleMpvCommandFromIpc(['add', 'sub-pos', 1], options);
  assert.deepEqual(sentCommands, [['add', 'sub-pos', 1]]);
  assert.deepEqual(osd, ['Subtitle position: ${sub-pos}']);
});

test('handleMpvCommandFromIpc emits osd for primary subtitle track keybinding proxies', () => {
  const { options, sentCommands, osd } = createOptions();
  handleMpvCommandFromIpc(['cycle', 'sid'], options);
  assert.deepEqual(sentCommands, [['cycle', 'sid']]);
  assert.deepEqual(osd, ['Subtitle track: ${sid}']);
});

test('handleMpvCommandFromIpc emits osd for secondary subtitle track keybinding proxies', () => {
  const { options, sentCommands, osd } = createOptions();
  handleMpvCommandFromIpc(['set_property', 'secondary-sid', 'auto'], options);
  assert.deepEqual(sentCommands, [['set_property', 'secondary-sid', 'auto']]);
  assert.deepEqual(osd, ['Secondary subtitle track: ${secondary-sid}']);
});

test('handleMpvCommandFromIpc does not forward commands while disconnected', () => {
  const { options, sentCommands, osd } = createOptions({
    isMpvConnected: () => false,
  });
  handleMpvCommandFromIpc(['add', 'sub-pos', 1], options);
  assert.deepEqual(sentCommands, []);
  assert.deepEqual(osd, []);
});
