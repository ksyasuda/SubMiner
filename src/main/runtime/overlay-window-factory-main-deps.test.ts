import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildCreateInvisibleWindowMainDepsHandler,
  createBuildCreateMainWindowMainDepsHandler,
  createBuildCreateModalWindowMainDepsHandler,
  createBuildCreateOverlayWindowMainDepsHandler,
  createBuildCreateSecondaryWindowMainDepsHandler,
} from './overlay-window-factory-main-deps';

test('overlay window factory main deps builders return mapped handlers', () => {
  const calls: string[] = [];
  const buildOverlayDeps = createBuildCreateOverlayWindowMainDepsHandler({
    createOverlayWindowCore: (kind) => ({ kind }),
    isDev: true,
    getOverlayDebugVisualizationEnabled: () => false,
    ensureOverlayWindowLevel: () => calls.push('ensure-level'),
    onRuntimeOptionsChanged: () => calls.push('runtime-options-changed'),
    setOverlayDebugVisualizationEnabled: (enabled) => calls.push(`debug:${enabled}`),
    isOverlayVisible: (kind) => kind === 'visible',
    tryHandleOverlayShortcutLocalFallback: () => false,
    onWindowClosed: (kind) => calls.push(`closed:${kind}`),
  });

  const overlayDeps = buildOverlayDeps();
  assert.equal(overlayDeps.isDev, true);
  assert.equal(overlayDeps.getOverlayDebugVisualizationEnabled(), false);
  assert.equal(overlayDeps.isOverlayVisible('visible'), true);

  const buildMainDeps = createBuildCreateMainWindowMainDepsHandler({
    createOverlayWindow: () => ({ id: 'visible' }),
    setMainWindow: () => calls.push('set-main'),
  });
  const mainDeps = buildMainDeps();
  mainDeps.setMainWindow(null);

  const buildInvisibleDeps = createBuildCreateInvisibleWindowMainDepsHandler({
    createOverlayWindow: () => ({ id: 'invisible' }),
    setInvisibleWindow: () => calls.push('set-invisible'),
  });
  const invisibleDeps = buildInvisibleDeps();
  invisibleDeps.setInvisibleWindow(null);

  const buildSecondaryDeps = createBuildCreateSecondaryWindowMainDepsHandler({
    createOverlayWindow: () => ({ id: 'secondary' }),
    setSecondaryWindow: () => calls.push('set-secondary'),
  });
  const secondaryDeps = buildSecondaryDeps();
  secondaryDeps.setSecondaryWindow(null);

  const buildModalDeps = createBuildCreateModalWindowMainDepsHandler({
    createOverlayWindow: () => ({ id: 'modal' }),
    setModalWindow: () => calls.push('set-modal'),
  });
  const modalDeps = buildModalDeps();
  modalDeps.setModalWindow(null);

  assert.deepEqual(calls, ['set-main', 'set-invisible', 'set-secondary', 'set-modal']);
});
