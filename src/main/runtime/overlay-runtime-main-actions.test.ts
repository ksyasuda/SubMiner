import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBroadcastRuntimeOptionsChangedHandler,
  createGetRuntimeOptionsStateHandler,
  createOpenRuntimeOptionsPaletteHandler,
  createRestorePreviousSecondarySubVisibilityHandler,
  createSendToActiveOverlayWindowHandler,
  createSetOverlayDebugVisualizationEnabledHandler,
} from './overlay-runtime-main-actions';

test('runtime options state handler returns empty list without manager', () => {
  const getState = createGetRuntimeOptionsStateHandler({
    getRuntimeOptionsManager: () => null,
  });
  assert.deepEqual(getState(), []);
});

test('runtime options state handler returns list from manager', () => {
  const getState = createGetRuntimeOptionsStateHandler({
    getRuntimeOptionsManager: () =>
      ({
        listOptions: () => [
          {
            id: 'anki.autoUpdateNewCards',
            label: 'X',
            scope: 'ankiConnect',
            valueType: 'boolean',
            value: true,
            allowedValues: [true, false],
            requiresRestart: false,
          },
        ],
      }) as never,
  });
  assert.deepEqual(getState(), [
    {
      id: 'anki.autoUpdateNewCards',
      label: 'X',
      scope: 'ankiConnect',
      valueType: 'boolean',
      value: true,
      allowedValues: [true, false],
      requiresRestart: false,
    },
  ]);
});

test('restore previous secondary subtitle visibility no-ops without connected mpv client', () => {
  let restored = false;
  const restore = createRestorePreviousSecondarySubVisibilityHandler({
    getMpvClient: () => ({
      connected: false,
      restorePreviousSecondarySubVisibility: () => (restored = true),
    }),
  });
  restore();
  assert.equal(restored, false);
});

test('restore previous secondary subtitle visibility calls runtime when connected', () => {
  let restored = false;
  const restore = createRestorePreviousSecondarySubVisibilityHandler({
    getMpvClient: () => ({
      connected: true,
      restorePreviousSecondarySubVisibility: () => (restored = true),
    }),
  });
  restore();
  assert.equal(restored, true);
});

test('broadcast runtime options changed passes through state getter and broadcaster', () => {
  const calls: string[] = [];
  const broadcast = createBroadcastRuntimeOptionsChangedHandler({
    broadcastRuntimeOptionsChangedRuntime: (getState, emit) => {
      calls.push(`state:${JSON.stringify(getState())}`);
      emit('runtime-options:changed', { id: 1 });
    },
    getRuntimeOptionsState: () => [
      {
        id: 'anki.autoUpdateNewCards',
        label: 'X',
        scope: 'ankiConnect',
        valueType: 'boolean',
        value: true,
        allowedValues: [true, false],
        requiresRestart: false,
      },
    ],
    broadcastToOverlayWindows: (channel, payload) =>
      calls.push(`emit:${channel}:${JSON.stringify(payload)}`),
  });

  broadcast();
  assert.deepEqual(calls, [
    'state:[{"id":"anki.autoUpdateNewCards","label":"X","scope":"ankiConnect","valueType":"boolean","value":true,"allowedValues":[true,false],"requiresRestart":false}]',
    'emit:runtime-options:changed:{"id":1}',
  ]);
});

test('send to active overlay window delegates to runtime sender', () => {
  const send = createSendToActiveOverlayWindowHandler({
    sendToActiveOverlayWindowRuntime: (channel, payload) => channel === 'ok' && payload === 1,
  });
  assert.equal(send('ok', 1), true);
  assert.equal(send('no', 1), false);
});

test('set overlay debug visualization enabled delegates with current state and broadcast', () => {
  const calls: string[] = [];
  let current = false;
  const setEnabled = createSetOverlayDebugVisualizationEnabledHandler({
    setOverlayDebugVisualizationEnabledRuntime: (curr, next, setCurrent) => {
      calls.push(`runtime:${curr}->${next}`);
      setCurrent(next);
      // no renderer-level side effects for this legacy debug path.
    },
    getCurrentEnabled: () => current,
    setCurrentEnabled: (enabled) => {
      current = enabled;
      calls.push(`set:${enabled}`);
    },
  });

  setEnabled(true);
  assert.equal(current, true);
  assert.deepEqual(calls, ['runtime:false->true', 'set:true']);
});

test('open runtime options palette handler delegates to runtime', () => {
  let opened = false;
  const open = createOpenRuntimeOptionsPaletteHandler({
    openRuntimeOptionsPaletteRuntime: () => {
      opened = true;
    },
  });
  open();
  assert.equal(opened, true);
});
