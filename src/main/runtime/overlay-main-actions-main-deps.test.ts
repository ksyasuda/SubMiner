import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildAppendClipboardVideoToQueueMainDepsHandler,
  createBuildHandleOverlayModalClosedMainDepsHandler,
  createBuildSetOverlayVisibleMainDepsHandler,
  createBuildToggleOverlayMainDepsHandler,
} from './overlay-main-actions-main-deps';

test('overlay main action main deps builders map callbacks', () => {
  const calls: string[] = [];

  const setOverlay = createBuildSetOverlayVisibleMainDepsHandler({
    setVisibleOverlayVisible: (visible) => calls.push(`set:${visible}`),
  })();
  setOverlay.setVisibleOverlayVisible(true);

  const toggleOverlay = createBuildToggleOverlayMainDepsHandler({
    toggleVisibleOverlay: () => calls.push('toggle'),
  })();
  toggleOverlay.toggleVisibleOverlay();

  const modalClosed = createBuildHandleOverlayModalClosedMainDepsHandler({
    handleOverlayModalClosedRuntime: (modal) => calls.push(`modal:${modal}`),
  })();
  modalClosed.handleOverlayModalClosedRuntime('runtime-options');

  const append = createBuildAppendClipboardVideoToQueueMainDepsHandler({
    appendClipboardVideoToQueueRuntime: () => {
      calls.push('append');
      return { ok: true, message: 'ok' };
    },
    getMpvClient: () => ({ connected: true }),
    readClipboardText: () => '/tmp/v.mkv',
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    sendMpvCommand: (command) => calls.push(`cmd:${command.join(':')}`),
  })();
  assert.deepEqual(
    append.appendClipboardVideoToQueueRuntime({
      getMpvClient: () => ({ connected: true }),
      readClipboardText: () => '/tmp/v.mkv',
      showMpvOsd: () => {},
      sendMpvCommand: () => {},
    }),
    { ok: true, message: 'ok' },
  );
  assert.equal(append.readClipboardText(), '/tmp/v.mkv');
  assert.equal(typeof append.getMpvClient(), 'object');
  append.showMpvOsd('queued');
  append.sendMpvCommand(['loadfile', '/tmp/v.mkv', 'append']);

  assert.deepEqual(calls, [
    'set:true',
    'toggle',
    'modal:runtime-options',
    'append',
    'osd:queued',
    'cmd:loadfile:/tmp/v.mkv:append',
  ]);
});
