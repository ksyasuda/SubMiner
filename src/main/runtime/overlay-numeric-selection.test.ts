import assert from 'node:assert/strict';
import test from 'node:test';

import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { tryBeginVisibleOverlayNumericSelection } from './overlay-numeric-selection';

function createWindowStub(
  options: {
    destroyed?: boolean;
    visible?: boolean;
    focused?: boolean;
    webContentsFocused?: boolean;
  } = {},
) {
  const calls: string[] = [];
  return {
    calls,
    window: {
      isDestroyed: () => options.destroyed === true,
      isVisible: () => options.visible !== false,
      isFocused: () => options.focused === true,
      setIgnoreMouseEvents: (ignore: boolean) => {
        calls.push(`mouse:${ignore}`);
      },
      focus: () => {
        calls.push('focus');
      },
      webContents: {
        isFocused: () => options.webContentsFocused === true,
        focus: () => {
          calls.push('web-focus');
        },
        send: (channel: string, payload: unknown) => {
          calls.push(`send:${channel}:${JSON.stringify(payload)}`);
        },
      },
    },
  };
}

test('tryBeginVisibleOverlayNumericSelection focuses visible overlay and sends selector event', () => {
  const { window, calls } = createWindowStub();

  const handled = tryBeginVisibleOverlayNumericSelection({
    actionId: 'copySubtitleMultiple',
    timeoutMs: 1234,
    getMainWindow: () => window,
    getVisibleOverlayVisible: () => true,
  });

  assert.equal(handled, true);
  assert.deepEqual(calls, [
    'mouse:false',
    'focus',
    'web-focus',
    `send:${IPC_CHANNELS.event.sessionNumericSelectionStart}:{"actionId":"copySubtitleMultiple","timeoutMs":1234}`,
  ]);
});

test('tryBeginVisibleOverlayNumericSelection skips hidden visible overlay', () => {
  const { window, calls } = createWindowStub({ visible: false });

  const handled = tryBeginVisibleOverlayNumericSelection({
    actionId: 'mineSentenceMultiple',
    timeoutMs: 3000,
    getMainWindow: () => window,
    getVisibleOverlayVisible: () => true,
  });

  assert.equal(handled, false);
  assert.deepEqual(calls, []);
});
