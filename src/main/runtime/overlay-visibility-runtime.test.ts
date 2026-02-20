import assert from 'node:assert/strict';
import test from 'node:test';
import { createOverlayVisibilityRuntime } from './overlay-visibility-runtime';

test('overlay visibility runtime wires set/toggle handlers through composed deps', () => {
  let visible = false;
  let invisible = true;
  let setVisibleCoreCalls = 0;
  let setInvisibleCoreCalls = 0;
  let lastBoundSubVisibility: boolean | null = null;

  const runtime = createOverlayVisibilityRuntime({
    setVisibleOverlayVisibleDeps: {
      setVisibleOverlayVisibleCore: (options) => {
        setVisibleCoreCalls += 1;
        options.setVisibleOverlayVisibleState(options.visible);
        options.updateVisibleOverlayVisibility();
        options.updateInvisibleOverlayVisibility();
        options.syncInvisibleOverlayMousePassthrough();
        if (options.shouldBindVisibleOverlayToMpvSubVisibility() && options.isMpvConnected()) {
          options.setMpvSubVisibility(options.visible);
        }
      },
      setVisibleOverlayVisibleState: (nextVisible) => {
        visible = nextVisible;
      },
      updateVisibleOverlayVisibility: () => {},
      updateInvisibleOverlayVisibility: () => {},
      syncInvisibleOverlayMousePassthrough: () => {},
      shouldBindVisibleOverlayToMpvSubVisibility: () => true,
      isMpvConnected: () => true,
      setMpvSubVisibility: (nextVisible) => {
        lastBoundSubVisibility = nextVisible;
      },
    },
    setInvisibleOverlayVisibleDeps: {
      setInvisibleOverlayVisibleCore: (options) => {
        setInvisibleCoreCalls += 1;
        options.setInvisibleOverlayVisibleState(options.visible);
        options.updateInvisibleOverlayVisibility();
        options.syncInvisibleOverlayMousePassthrough();
      },
      setInvisibleOverlayVisibleState: (nextVisible) => {
        invisible = nextVisible;
      },
      updateInvisibleOverlayVisibility: () => {},
      syncInvisibleOverlayMousePassthrough: () => {},
    },
    getVisibleOverlayVisible: () => visible,
    getInvisibleOverlayVisible: () => invisible,
  });

  runtime.setVisibleOverlayVisible(true);
  assert.equal(visible, true);
  assert.equal(lastBoundSubVisibility, true);

  runtime.toggleVisibleOverlay();
  assert.equal(visible, false);

  runtime.setOverlayVisible(true);
  assert.equal(visible, true);

  runtime.toggleOverlay();
  assert.equal(visible, false);

  runtime.setInvisibleOverlayVisible(false);
  assert.equal(invisible, false);

  runtime.toggleInvisibleOverlay();
  assert.equal(invisible, true);

  assert.equal(setVisibleCoreCalls, 4);
  assert.equal(setInvisibleCoreCalls, 2);
});
