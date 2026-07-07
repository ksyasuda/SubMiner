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
    setVisibleOverlayVisible: (visible) => calls.push(`visible:${visible}`),
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
  assert.equal(deps.getResolver(), resolver);
  assert.equal(deps.getRestoreVisibleOverlayOnModalClose(), modalSet);
  deps.setVisibleOverlayVisible(true);
  deps.setResolver(null);
  assert.equal(deps.sendToVisibleOverlay('kiku:open', 1), true);
  assert.deepEqual(calls, ['visible:true', 'set-resolver:null', 'send:kiku:open:1']);
});

test('field grouping overlay main deps builder forwards modal open/teardown/prereq wiring', () => {
  // Regression: these are optional on the runtime options, so a missing forward compiled
  // silently and left the field grouping modal with no ack/retry, no teardown, and no prereqs —
  // the reason the earlier recovery fixes never took effect.
  const calls: string[] = [];
  const waitForModalOpen = async (modal: 'kiku', timeoutMs: number): Promise<boolean> => {
    calls.push(`wait:${modal}:${timeoutMs}`);
    return true;
  };
  const handleOverlayModalClosed = (modal: 'kiku'): void => {
    calls.push(`closed:${modal}`);
  };
  const logWarn = (message: string): void => {
    calls.push(`warn:${message}`);
  };
  const ensureOverlayStartupPrereqs = (): void => {
    calls.push('prereqs');
  };
  const ensureOverlayWindowsReadyForVisibilityActions = (): void => {
    calls.push('windows-ready');
  };

  const deps = createBuildFieldGroupingOverlayMainDepsHandler<'kiku'>({
    getMainWindow: () => null,
    getVisibleOverlayVisible: () => false,
    setVisibleOverlayVisible: () => {},
    getResolver: () => null,
    setResolver: () => {},
    getRestoreVisibleOverlayOnModalClose: () => new Set<'kiku'>(),
    waitForModalOpen,
    handleOverlayModalClosed,
    logWarn,
    ensureOverlayStartupPrereqs,
    ensureOverlayWindowsReadyForVisibilityActions,
    sendToActiveOverlayWindow: () => true,
  })();

  assert.equal(deps.waitForModalOpen, waitForModalOpen);
  assert.equal(deps.handleOverlayModalClosed, handleOverlayModalClosed);
  assert.equal(deps.logWarn, logWarn);
  assert.equal(deps.ensureOverlayStartupPrereqs, ensureOverlayStartupPrereqs);
  assert.equal(
    deps.ensureOverlayWindowsReadyForVisibilityActions,
    ensureOverlayWindowsReadyForVisibilityActions,
  );
});
