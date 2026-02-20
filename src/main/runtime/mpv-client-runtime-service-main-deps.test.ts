import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildMpvClientRuntimeServiceFactoryDepsHandler } from './mpv-client-runtime-service-main-deps';

test('mpv runtime service main deps builder maps state and callbacks', () => {
  const calls: string[] = [];
  let reconnectTimer: ReturnType<typeof setTimeout> | null = null;

  class FakeClient {
    constructor(public socketPath: string, public options: unknown) {}
  }

  const build = createBuildMpvClientRuntimeServiceFactoryDepsHandler({
    createClient: FakeClient,
    getSocketPath: () => '/tmp/mpv.sock',
    getResolvedConfig: () => ({ mode: 'test' }),
    isAutoStartOverlayEnabled: () => true,
    setOverlayVisible: (visible) => calls.push(`overlay:${visible}`),
    shouldBindVisibleOverlayToMpvSubVisibility: () => true,
    isVisibleOverlayVisible: () => false,
    getReconnectTimer: () => reconnectTimer,
    setReconnectTimer: (timer) => {
      reconnectTimer = timer;
      calls.push('set-reconnect');
    },
    bindEventHandlers: () => calls.push('bind'),
  });

  const deps = build();
  assert.equal(deps.socketPath, '/tmp/mpv.sock');
  assert.equal(deps.options.autoStartOverlay, true);
  assert.equal(deps.options.shouldBindVisibleOverlayToMpvSubVisibility(), true);
  assert.equal(deps.options.isVisibleOverlayVisible(), false);
  assert.deepEqual(deps.options.getResolvedConfig(), { mode: 'test' });

  deps.options.setOverlayVisible(true);
  deps.options.setReconnectTimer(setTimeout(() => {}, 0));
  deps.bindEventHandlers(new FakeClient('/tmp/mpv.sock', {}));

  assert.ok(calls.includes('overlay:true'));
  assert.ok(calls.includes('set-reconnect'));
  assert.ok(calls.includes('bind'));
  assert.ok(reconnectTimer);
});
