import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createAppendClipboardVideoToQueueHandler,
  createHandleOverlayModalClosedHandler,
  createSetOverlayVisibleHandler,
  createToggleOverlayHandler,
} from './overlay-main-actions';

test('set overlay visible handler delegates to visible overlay setter', () => {
  const calls: string[] = [];
  const setOverlayVisible = createSetOverlayVisibleHandler({
    setVisibleOverlayVisible: (visible) => calls.push(`set:${visible}`),
  });

  setOverlayVisible(true);
  assert.deepEqual(calls, ['set:true']);
});

test('toggle overlay handler delegates to visible toggle', () => {
  const calls: string[] = [];
  const toggleOverlay = createToggleOverlayHandler({
    toggleVisibleOverlay: () => calls.push('toggle'),
  });

  toggleOverlay();
  assert.deepEqual(calls, ['toggle']);
});

test('overlay modal closed handler delegates to runtime handler', () => {
  const calls: string[] = [];
  const handleClosed = createHandleOverlayModalClosedHandler({
    handleOverlayModalClosedRuntime: (modal) => calls.push(`closed:${modal}`),
  });

  handleClosed('runtime-options');
  assert.deepEqual(calls, ['closed:runtime-options']);
});

test('append clipboard queue handler forwards runtime deps and result', () => {
  const calls: string[] = [];
  const mpvClient = { connected: true };
  const appendClipboardVideoToQueue = createAppendClipboardVideoToQueueHandler({
    appendClipboardVideoToQueueRuntime: (options) => {
      assert.equal(options.getMpvClient(), mpvClient);
      assert.equal(options.readClipboardText(), '/tmp/video.mkv');
      options.showMpvOsd('queued');
      options.sendMpvCommand(['loadfile', '/tmp/video.mkv', 'append']);
      return { ok: true, message: 'ok' };
    },
    getMpvClient: () => mpvClient,
    readClipboardText: () => '/tmp/video.mkv',
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    sendMpvCommand: (command) => calls.push(`mpv:${command.join(':')}`),
  });

  const result = appendClipboardVideoToQueue();
  assert.deepEqual(result, { ok: true, message: 'ok' });
  assert.deepEqual(calls, ['osd:queued', 'mpv:loadfile:/tmp/video.mkv:append']);
});
