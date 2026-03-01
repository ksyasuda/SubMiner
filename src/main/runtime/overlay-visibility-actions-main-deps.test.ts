import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildSetVisibleOverlayVisibleMainDepsHandler,
  createBuildToggleVisibleOverlayMainDepsHandler,
} from './overlay-visibility-actions-main-deps';

test('overlay visibility action main deps builders map callbacks', () => {
  const calls: string[] = [];
  let warmupStarts = 0;

  const setVisible = createBuildSetVisibleOverlayVisibleMainDepsHandler({
    setVisibleOverlayVisibleCore: () => calls.push('visible-core'),
    setVisibleOverlayVisibleState: (visible) => calls.push(`visible-state:${visible}`),
    updateVisibleOverlayVisibility: () => calls.push('update-visible'),
    onVisibleOverlayEnabled: () => {
      warmupStarts += 1;
    },
  })();
  setVisible.setVisibleOverlayVisibleCore({
    visible: true,
    setVisibleOverlayVisibleState: () => {},
    updateVisibleOverlayVisibility: () => {},
  });
  setVisible.setVisibleOverlayVisibleState(true);
  setVisible.updateVisibleOverlayVisibility();
  setVisible.onVisibleOverlayEnabled?.();

  const toggleVisible = createBuildToggleVisibleOverlayMainDepsHandler({
    getVisibleOverlayVisible: () => false,
    setVisibleOverlayVisible: (visible) => calls.push(`toggle-visible:${visible}`),
  })();
  assert.equal(toggleVisible.getVisibleOverlayVisible(), false);
  toggleVisible.setVisibleOverlayVisible(true);

  assert.deepEqual(calls, [
    'visible-core',
    'visible-state:true',
    'update-visible',
    'toggle-visible:true',
  ]);
  assert.equal(warmupStarts, 1);
});
