import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildCreateMainWindowMainDepsHandler,
  createBuildCreateModalWindowMainDepsHandler,
  createBuildCreateOverlayWindowMainDepsHandler,
} from './overlay-window-factory-main-deps';

test('overlay window factory main deps builders return mapped handlers', () => {
  const calls: string[] = [];
  const yomitanSession = { id: 'session' } as never;
  const buildOverlayDeps = createBuildCreateOverlayWindowMainDepsHandler({
    createOverlayWindowCore: (kind) => ({ kind }),
    isDev: true,
    ensureOverlayWindowLevel: () => calls.push('ensure-level'),
    onRuntimeOptionsChanged: () => calls.push('runtime-options-changed'),
    setOverlayDebugVisualizationEnabled: (enabled) => calls.push(`debug:${enabled}`),
    isOverlayVisible: (kind) => kind === 'visible',
    tryHandleOverlayShortcutLocalFallback: () => false,
    forwardTabToMpv: () => calls.push('forward-tab'),
    onVisibleWindowFocused: () => calls.push('visible-focus'),
    onWindowClosed: (kind) => calls.push(`closed:${kind}`),
    getYomitanSession: () => yomitanSession,
  });

  const overlayDeps = buildOverlayDeps();
  assert.equal(overlayDeps.isDev, true);
  assert.equal(overlayDeps.isOverlayVisible('visible'), true);
  assert.equal(overlayDeps.getYomitanSession(), yomitanSession);
  overlayDeps.forwardTabToMpv();
  overlayDeps.onVisibleWindowFocused?.();

  const buildMainDeps = createBuildCreateMainWindowMainDepsHandler({
    getMainWindow: () => null,
    isWindowDestroyed: () => false,
    createOverlayWindow: () => ({ id: 'visible' }),
    setMainWindow: () => calls.push('set-main'),
  });
  const mainDeps = buildMainDeps();
  assert.equal(mainDeps.getMainWindow(), null);
  assert.equal(mainDeps.isWindowDestroyed({ id: 'visible' }), false);
  mainDeps.setMainWindow(null);

  const buildModalDeps = createBuildCreateModalWindowMainDepsHandler({
    createOverlayWindow: () => ({ id: 'modal' }),
    setModalWindow: () => calls.push('set-modal'),
  });
  const modalDeps = buildModalDeps();
  modalDeps.setModalWindow(null);

  assert.deepEqual(calls, ['forward-tab', 'visible-focus', 'set-main', 'set-modal']);
});
