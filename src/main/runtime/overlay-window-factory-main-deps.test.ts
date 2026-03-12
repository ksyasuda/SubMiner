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
    onWindowClosed: (kind) => calls.push(`closed:${kind}`),
    getYomitanSession: () => yomitanSession,
  });

  const overlayDeps = buildOverlayDeps();
  assert.equal(overlayDeps.isDev, true);
  assert.equal(overlayDeps.isOverlayVisible('visible'), true);
  assert.equal(overlayDeps.getYomitanSession(), yomitanSession);
  overlayDeps.forwardTabToMpv();

  const buildMainDeps = createBuildCreateMainWindowMainDepsHandler({
    createOverlayWindow: () => ({ id: 'visible' }),
    setMainWindow: () => calls.push('set-main'),
  });
  const mainDeps = buildMainDeps();
  mainDeps.setMainWindow(null);

  const buildModalDeps = createBuildCreateModalWindowMainDepsHandler({
    createOverlayWindow: () => ({ id: 'modal' }),
    setModalWindow: () => calls.push('set-modal'),
  });
  const modalDeps = buildModalDeps();
  modalDeps.setModalWindow(null);

  assert.deepEqual(calls, ['forward-tab', 'set-main', 'set-modal']);
});
