import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildSetInvisibleOverlayVisibleMainDepsHandler,
  createBuildSetVisibleOverlayVisibleMainDepsHandler,
  createBuildToggleInvisibleOverlayMainDepsHandler,
  createBuildToggleVisibleOverlayMainDepsHandler,
} from './overlay-visibility-actions-main-deps';

test('overlay visibility action main deps builders map callbacks', () => {
  const calls: string[] = [];

  const setVisible = createBuildSetVisibleOverlayVisibleMainDepsHandler({
    setVisibleOverlayVisibleCore: () => calls.push('visible-core'),
    setVisibleOverlayVisibleState: (visible) => calls.push(`visible-state:${visible}`),
    updateVisibleOverlayVisibility: () => calls.push('update-visible'),
    updateInvisibleOverlayVisibility: () => calls.push('update-invisible'),
    syncInvisibleOverlayMousePassthrough: () => calls.push('sync'),
    shouldBindVisibleOverlayToMpvSubVisibility: () => true,
    isMpvConnected: () => true,
    setMpvSubVisibility: (visible) => calls.push(`mpv:${visible}`),
  })();
  setVisible.setVisibleOverlayVisibleCore({
    visible: true,
    setVisibleOverlayVisibleState: () => {},
    updateVisibleOverlayVisibility: () => {},
    updateInvisibleOverlayVisibility: () => {},
    syncInvisibleOverlayMousePassthrough: () => {},
    shouldBindVisibleOverlayToMpvSubVisibility: () => true,
    isMpvConnected: () => true,
    setMpvSubVisibility: () => {},
  });
  setVisible.setVisibleOverlayVisibleState(true);
  setVisible.updateVisibleOverlayVisibility();
  setVisible.updateInvisibleOverlayVisibility();
  setVisible.syncInvisibleOverlayMousePassthrough();
  assert.equal(setVisible.shouldBindVisibleOverlayToMpvSubVisibility(), true);
  assert.equal(setVisible.isMpvConnected(), true);
  setVisible.setMpvSubVisibility(false);

  const setInvisible = createBuildSetInvisibleOverlayVisibleMainDepsHandler({
    setInvisibleOverlayVisibleCore: () => calls.push('invisible-core'),
    setInvisibleOverlayVisibleState: (visible) => calls.push(`invisible-state:${visible}`),
    updateInvisibleOverlayVisibility: () => calls.push('update-only-invisible'),
    syncInvisibleOverlayMousePassthrough: () => calls.push('sync-only'),
  })();
  setInvisible.setInvisibleOverlayVisibleCore({
    visible: false,
    setInvisibleOverlayVisibleState: () => {},
    updateInvisibleOverlayVisibility: () => {},
    syncInvisibleOverlayMousePassthrough: () => {},
  });
  setInvisible.setInvisibleOverlayVisibleState(false);
  setInvisible.updateInvisibleOverlayVisibility();
  setInvisible.syncInvisibleOverlayMousePassthrough();

  const toggleVisible = createBuildToggleVisibleOverlayMainDepsHandler({
    getVisibleOverlayVisible: () => false,
    setVisibleOverlayVisible: (visible) => calls.push(`toggle-visible:${visible}`),
  })();
  assert.equal(toggleVisible.getVisibleOverlayVisible(), false);
  toggleVisible.setVisibleOverlayVisible(true);

  const toggleInvisible = createBuildToggleInvisibleOverlayMainDepsHandler({
    getInvisibleOverlayVisible: () => true,
    setInvisibleOverlayVisible: (visible) => calls.push(`toggle-invisible:${visible}`),
  })();
  assert.equal(toggleInvisible.getInvisibleOverlayVisible(), true);
  toggleInvisible.setInvisibleOverlayVisible(false);

  assert.deepEqual(calls, [
    'visible-core',
    'visible-state:true',
    'update-visible',
    'update-invisible',
    'sync',
    'mpv:false',
    'invisible-core',
    'invisible-state:false',
    'update-only-invisible',
    'sync-only',
    'toggle-visible:true',
    'toggle-invisible:false',
  ]);
});
