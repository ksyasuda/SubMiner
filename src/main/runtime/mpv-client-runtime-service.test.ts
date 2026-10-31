import assert from 'node:assert/strict';
import test from 'node:test';
import { DEFAULT_CONFIG } from '../../config';
import {
  createMpvClientRuntimeServiceFactory,
  type MpvClientRuntimeServiceOptions,
} from './mpv-client-runtime-service';

test('mpv client factory reads socket path and auto-start per client, binds before connect', () => {
  const calls: string[] = [];
  let socketPath = '/tmp/a.sock';
  let autoStart = false;

  class FakeClient {
    constructor(
      readonly socketPath: string,
      readonly options: MpvClientRuntimeServiceOptions,
    ) {}
    connect() {
      calls.push(`connect:${this.socketPath}`);
    }
  }

  const createClient = createMpvClientRuntimeServiceFactory({
    createClient: FakeClient,
    getSocketPath: () => socketPath,
    getResolvedConfig: () => DEFAULT_CONFIG,
    isAutoStartOverlayEnabled: () => autoStart,
    setOverlayVisible: () => {},
    isVisibleOverlayVisible: () => false,
    getReconnectTimer: () => null,
    setReconnectTimer: () => {},
    bindEventHandlers: (client) => calls.push(`bind:${client.socketPath}`),
  });

  const first = createClient();
  socketPath = '/tmp/b.sock';
  autoStart = true;
  const second = createClient();

  assert.deepEqual(calls, [
    'bind:/tmp/a.sock',
    'connect:/tmp/a.sock',
    'bind:/tmp/b.sock',
    'connect:/tmp/b.sock',
  ]);
  assert.equal(first.options.autoStartOverlay, false);
  assert.equal(second.options.autoStartOverlay, true);
  assert.equal(first.options.shouldQuitOnMpvShutdown?.(), false);
});
