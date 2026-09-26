import test from 'node:test';
import assert from 'node:assert/strict';
import { buildHachidoriSharingScript, parseHachidoriHostStatus } from './hachidori-sharing';
import { runInNewContext } from 'node:vm';

test('host status rejects malformed replies and distinguishes local, connected and offline', () => {
  assert.deepEqual(parseHachidoriHostStatus({ ok: true, sharing: { client: { linked: false } } }), {
    kind: 'local',
  });
  assert.throws(
    () => parseHachidoriHostStatus({ ok: false, error: 'Connection refused' }),
    /Connection refused/,
  );
  assert.throws(
    () =>
      parseHachidoriHostStatus({
        ok: true,
        sharing: {
          client: { linked: true, connected: true, address: 'host', host: { dictionaryCount: -1 } },
        },
      }),
    /dictionary count/,
  );
  assert.equal(
    parseHachidoriHostStatus({
      ok: true,
      sharing: { client: { linked: true, address: 'host', connected: false } },
    }).kind,
    'disconnected',
  );
});

test('sharing script safely links an address and checks live engine inventory', async () => {
  const address = `ws://host:8771/link?quote="`;
  const requests: Array<Record<string, unknown>> = [];
  const script = buildHachidoriSharingScript({ type: 'hd_sharing_client_link', address });
  const value: unknown = await runInNewContext(script, {
    crypto,
    setTimeout,
    clearTimeout,
    chrome: {
      runtime: {
        sendMessage: async (message: Record<string, unknown>) => {
          requests.push(message);
          if (message.type === 'hd_status') return { ok: true, ready: true, loading: false };
          if (message.type === 'hd_state_read')
            return { ok: true, state: { dictionaries: [{}, {}] } };
          return {
            ok: true,
            sharing: {
              client: {
                linked: true,
                connected: true,
                address,
                host: { name: 'Host', dictionaryCount: 99 },
              },
            },
          };
        },
      },
    },
  });
  assert.equal(requests[0]?.address, address);
  assert.deepEqual(parseHachidoriHostStatus(value), {
    kind: 'connected',
    address,
    name: 'Host',
    dictionaryCount: 2,
  });
});

test('an unresponsive host preserves its address so setup can still unlink', async () => {
  const value: unknown = await runInNewContext(
    buildHachidoriSharingScript({ type: 'hd_sharing_status' }),
    {
      crypto,
      setTimeout,
      clearTimeout,
      chrome: {
        runtime: {
          sendMessage: async (message: { type: string }) => {
            if (message.type === 'hd_status') throw new Error('Host stopped responding');
            return {
              ok: true,
              sharing: {
                client: {
                  linked: true,
                  connected: true,
                  address: 'ws://host:8771/link',
                  host: { dictionaryCount: 1 },
                },
              },
            };
          },
        },
      },
    },
  );
  assert.deepEqual(parseHachidoriHostStatus(value), {
    kind: 'disconnected',
    address: 'ws://host:8771/link',
    message: 'Host stopped responding',
  });
});

for (const stalledCall of ['initial', 'refresh']) {
  test(`sharing ${stalledCall} requests time out and clear their timers`, async () => {
    const timers = new Map<number, () => void>();
    let nextTimer = 0;
    let sharingCalls = 0;
    let stalled = false;
    const result: Promise<unknown> = runInNewContext(
      buildHachidoriSharingScript({ type: 'hd_sharing_status' }),
      {
        crypto,
        setTimeout: (callback: () => void, delay: number) => {
          assert.equal(delay, 5000);
          timers.set(++nextTimer, callback);
          return nextTimer;
        },
        clearTimeout: (id: number) => timers.delete(id),
        chrome: {
          runtime: {
            sendMessage: async (message: { type: string }) => {
              if (message.type === 'hd_status') return { ok: true, ready: true };
              sharingCalls += 1;
              if (sharingCalls === (stalledCall === 'initial' ? 1 : 2)) {
                stalled = true;
                return new Promise(() => {});
              }
              return {
                ok: true,
                sharing: {
                  client: {
                    linked: true,
                    connected: true,
                    address: 'host',
                    host: { dictionaryCount: 1 },
                  },
                },
              };
            },
          },
        },
      },
    );
    for (let i = 0; i < 20 && !stalled; i++) await Promise.resolve();
    assert.equal(stalled, true);
    assert.equal(timers.size, 1);
    for (const callback of timers.values()) callback();
    if (stalledCall === 'initial') {
      await assert.rejects(result, /host did not respond/);
    } else {
      const status = parseHachidoriHostStatus(await result);
      assert.equal(status.kind, 'disconnected');
      if (status.kind === 'disconnected') assert.match(status.message, /host did not respond/);
    }
    assert.equal(timers.size, 0);
  });
}
