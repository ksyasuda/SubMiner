import assert from 'node:assert/strict';
import test from 'node:test';
import { createOverlayVisibilityRuntime } from './overlay-visibility-runtime';

test('overlay visibility runtime wires set/toggle handlers through composed deps', () => {
  let visible = false;
  let setVisibleCoreCalls = 0;
  let warmupStarts = 0;

  const runtime = createOverlayVisibilityRuntime({
    setVisibleOverlayVisibleDeps: {
      setVisibleOverlayVisibleCore: (options) => {
        setVisibleCoreCalls += 1;
        options.setVisibleOverlayVisibleState(options.visible);
        options.updateVisibleOverlayVisibility();
      },
      setVisibleOverlayVisibleState: (nextVisible) => {
        visible = nextVisible;
      },
      updateVisibleOverlayVisibility: () => {},
      onVisibleOverlayEnabled: () => {
        warmupStarts += 1;
      },
    },
    getVisibleOverlayVisible: () => visible,
  });

  runtime.setVisibleOverlayVisible(true);
  assert.equal(visible, true);
  runtime.setVisibleOverlayVisible(true);
  assert.equal(setVisibleCoreCalls, 1);

  runtime.toggleVisibleOverlay();
  assert.equal(visible, false);

  runtime.setOverlayVisible(true);
  assert.equal(visible, true);

  runtime.toggleOverlay();
  assert.equal(visible, false);

  assert.equal(setVisibleCoreCalls, 4);
  assert.equal(warmupStarts, 2);
});
