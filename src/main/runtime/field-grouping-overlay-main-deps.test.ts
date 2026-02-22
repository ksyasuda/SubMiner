import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildFieldGroupingOverlayMainDepsHandler } from './field-grouping-overlay-main-deps';

test('field grouping overlay main deps builder maps window visibility and resolver wiring', () => {
  const calls: string[] = [];
  const modalSet = new Set<'runtime-options'>();
  const resolver = (choice: unknown) => calls.push(`resolver:${choice}`);

  const deps = createBuildFieldGroupingOverlayMainDepsHandler({
    getMainWindow: () => ({
      isDestroyed: () => false,
      webContents: {
        send: () => {},
      },
    }),
    getVisibleOverlayVisible: () => true,
    getInvisibleOverlayVisible: () => false,
    setVisibleOverlayVisible: (visible) => calls.push(`visible:${visible}`),
    setInvisibleOverlayVisible: (visible) => calls.push(`invisible:${visible}`),
    getResolver: () => resolver,
    setResolver: (nextResolver) => {
      calls.push(`set-resolver:${nextResolver ? 'set' : 'null'}`);
    },
    getRestoreVisibleOverlayOnModalClose: () => modalSet,
    sendToActiveOverlayWindow: (channel, payload) => {
      calls.push(`send:${channel}:${String(payload)}`);
      return true;
    },
  })();

  assert.equal(deps.getMainWindow()?.isDestroyed(), false);
  assert.equal(deps.getVisibleOverlayVisible(), true);
  assert.equal(deps.getInvisibleOverlayVisible(), false);
  assert.equal(deps.getResolver(), resolver);
  assert.equal(deps.getRestoreVisibleOverlayOnModalClose(), modalSet);
  deps.setVisibleOverlayVisible(true);
  deps.setInvisibleOverlayVisible(false);
  deps.setResolver(null);
  assert.equal(deps.sendToVisibleOverlay('kiku:open', 1), true);
  assert.deepEqual(calls, [
    'visible:true',
    'invisible:false',
    'set-resolver:null',
    'send:kiku:open:1',
  ]);
});
