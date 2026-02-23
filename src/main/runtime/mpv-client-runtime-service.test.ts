import test from 'node:test';
import assert from 'node:assert/strict';
import { createMpvClientRuntimeServiceFactory } from './mpv-client-runtime-service';

test('mpv runtime service factory constructs client, binds handlers, and connects', () => {
  const calls: string[] = [];
  let constructedSocketPath = '';

  class FakeClient {
    connect(): void {
      calls.push('connect');
    }
    constructor(socketPath: string) {
      constructedSocketPath = socketPath;
      calls.push('construct');
    }
  }

  const createRuntimeService = createMpvClientRuntimeServiceFactory({
    createClient: FakeClient,
    socketPath: '/tmp/mpv.sock',
    options: {
      getResolvedConfig: () => ({}),
      autoStartOverlay: true,
      setOverlayVisible: () => {},
      shouldBindVisibleOverlayToMpvSubVisibility: () => false,
      isVisibleOverlayVisible: () => false,
      getReconnectTimer: () => null,
      setReconnectTimer: () => {},
    },
    bindEventHandlers: () => {
      calls.push('bind');
    },
  });

  const client = createRuntimeService();
  assert.ok(client instanceof FakeClient);
  assert.equal(constructedSocketPath, '/tmp/mpv.sock');
  assert.deepEqual(calls, ['construct', 'bind', 'connect']);
});
